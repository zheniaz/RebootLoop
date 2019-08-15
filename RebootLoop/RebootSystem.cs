using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using IWshRuntimeLibrary;
using System.Threading;
using System.Security.AccessControl;

namespace RebootLoop
{
    class RebootSystem
    {

        public string AppPath { get { return Path.GetDirectoryName(Path.GetFullPath(Constants.AppName)); } set { } }
        public string AppFullPath { get { return this.AppPath + "\\" + Constants.AppName; } set { } }
        public string LogFileFullPath { get { return IsLogFileExists ? Path.GetFullPath($"{AppPath}\\{Constants.LogFileName}") : ""; } }
        public string StartupDirectoryFullPath { get { return GetStartupPath(); } set { } }
        public string ShortcutRebootLoopFullPath { get { return $"{StartupDirectoryFullPath}\\{Constants.RebootLoopShortcutName}"; } }

        public int NeedToReboot { get { return GetCountNeedToReboot(); } set { } }
        public int RebootCount = 0;

        public bool IsLogFileExists { get { return CheckIfFileExists($"{AppPath}\\{Constants.LogFileName}"); } set { } }
        public bool IsShortcutOfRebootLoopExists { get { return CheckIfFileExists($"{StartupDirectoryFullPath}\\{Constants.RebootLoopShortcutName}.lnk"); } }

        public RebootSystem()
        {
            this.RebootCount = GetRebootCountFromLogFile();
        }

        public void CheckIfWantToReboot()
        {
            bool isWantToReboot = CheckYesOrNo("Do you want to reboot system?");
            if (isWantToReboot)
            {
                int rebootCount = CheckHowManyTimesReboot();
                if (rebootCount > 0 && rebootCount <= 10)
                {
                    CreateLogFile(rebootCount);
                    if (!IsShortcutOfRebootLoopExists)
                    {
                        CreateShortcut(AppPath, Constants.RebootLoopShortcutName);
                        MoveFileToFolder(AppPath + "\\" + Constants.RebootLoopShortcutName + ".lnk", StartupDirectoryFullPath + "\\" + Constants.RebootLoopShortcutName + ".lnk");
                    }
                    Reboot();
                }
                else if (rebootCount == 0)
                {
                    return;
                }
            }
        }

        public bool CheckYesOrNo(string query)
        {
            Console.WriteLine(query);
            Console.WriteLine("Press Y/N");
            char c = Console.ReadKey().KeyChar;
            Console.WriteLine();

            while (c != 'y' && c != 'Y' && c != 'n' && c != 'N')
            {
                Console.WriteLine("Press Y/N");
                c = Console.ReadKey().KeyChar;
                Console.WriteLine();
            }
            return (c == 'y' || c == 'Y') ? true : false;
        }

        public int CheckHowManyTimesReboot()
        {
            Console.WriteLine("How many times do you want to reboot your system?");
            Console.WriteLine("Enter digit from 1 to 10 or press n to quit:");
            string rebootCountString = Console.ReadLine();
            if (rebootCountString == "n") return 0;
            int rebootCount = 0;
            if (!Int32.TryParse(rebootCountString, out rebootCount) || rebootCount <= 0 || rebootCount > 10)
            {
                Console.WriteLine("Enter digit from 1 to 10 or press n to quit:");
            }
            while (rebootCount < 1 || rebootCount > 10)
            {
                rebootCountString = Console.ReadLine();
                if (rebootCountString == "n") return 0;
                if (!Int32.TryParse(rebootCountString, out rebootCount) || rebootCount <= 0 || rebootCount > 10)
                {
                    Console.WriteLine("Enter digit from 1 to 10 or press n to quit:");
                }
            }
            return rebootCount;
        }

        public void Reboot()
        {
            if (RebootCount > 0 && RebootCount <= NeedToReboot)
            {
                ReadAndRewriteRebootCountInLogFile();
                AddEntryToLogFile();
            }
            if (NeedToReboot == GetRebootsCountFromLogFile())
            {
                FinishRebooting();
                return;
            }
            string time = (NeedToReboot - RebootCount == 1) ? "time" : "times";
            Console.WriteLine($"{NeedToReboot - RebootCount} {time} left to reboot the system");
            Thread.Sleep(1500);
            Console.WriteLine("Rebooting System...");
            SetAutoLogOn();
            Thread.Sleep(2500);
            System.Diagnostics.Process.Start("ShutDown", "-r -t 0");
        }

        private void FinishRebooting()
        {
            RenameLogFile();
            Console.WriteLine("Removing ShortCut from Sturtup folder");
            Thread.Sleep(3000);
            RemoveFile(StartupDirectoryFullPath, Constants.RebootLoopShortcutName + ".lnk");
        }

        #region Region For Working  With Files

        public bool CheckIfFileExists(string filePath)
        {
            bool isFileExists = System.IO.File.Exists(filePath);
            return isFileExists;
        }

        public void CreateLogFile(int rebootCount)
        {
            string path = $"{AppPath}\\{Constants.LogFileName}";
            if (!System.IO.File.Exists(path))
            {
                using (StreamWriter sw = System.IO.File.CreateText(path))
                {
                    sw.WriteLine($"Created: {DateTime.Now}");
                    sw.WriteLine($"Need to reboot #{rebootCount}# times");
                    sw.WriteLine("rebooted: |0| times");
                }
            }
        }

        public void ReadAndRewriteRebootCountInLogFile()
        {
            if (IsLogFileExists == false)
            {
                throw new Exception("File rebootLog.txt does not exist.");
            }

            try
            {
                string text = System.IO.File.ReadAllText($"{ Constants.LogFileName}"); //"rebooted: |0| times";
                string rebootCountString = Regex.Match(text, @"(?<=\|)(.*?)(?=\|)").ToString();
                int rebootCount = 0;
                if (!Int32.TryParse(rebootCountString, out rebootCount))
                {
                    throw new Exception("Cannot convert from string to int.");
                }
                text = text.Replace($@"|{rebootCount}|", $@"|{++rebootCount}|");

                System.IO.File.WriteAllText(LogFileFullPath, text);
            }
            catch (Exception)
            {
                throw new Exception("Something went wrong!");
            }
        }

        public void AddEntryToLogFile()
        {
            string path = $"{AppPath}\\{Constants.LogFileName}";
            if (System.IO.File.Exists(path))
            {
                using (StreamWriter sw = System.IO.File.AppendText(path))
                {
                    sw.WriteLine("");
                    sw.WriteLine($"reboot #{RebootCount}");
                    sw.WriteLine($"rebooted: {DateTime.Now}");
                }
            }
        }

        public int GetRebootsCountFromLogFile()
        {
            string text = System.IO.File.ReadAllText($"{Constants.LogFileName}");
            string rebootCountString = Regex.Match(text, @"(?<=\|)(.*?)(?=\|)").ToString();
            int rebooted = 0;
            Int32.TryParse(rebootCountString, out rebooted);
            return rebooted;
        }

        private int GetRebootCountFromLogFile()
        {
            if (!IsLogFileExists) return 0;

            string text = System.IO.File.ReadAllText($"{Constants.LogFileName}");
            string rebootCountString = Regex.Match(text, @"(?<=\|)(.*?)(?=\|)").ToString();
            int rebooted = 0;
            if (!Int32.TryParse(rebootCountString, out rebooted))
            {
                throw new Exception("Cannot convert from string to int.");
            }
            return ++rebooted;
        }

        private int GetCountNeedToReboot()
        {
            string text = System.IO.File.ReadAllText($"{Constants.LogFileName}");
            string rebootCountString = Regex.Match(text, @"(?<=\#)(.*?)(?=\#)").ToString();
            int needToReboot = 0;
            if (!Int32.TryParse(rebootCountString, out needToReboot))
            {
                throw new Exception("Cannot convert from string to int.");
            }
            return needToReboot;
        }

        private void MoveFileToFolder(string filePath, string destination)
        {
            if (CheckIfFileExists(filePath))
            {
                System.IO.File.Move(filePath, destination);
            }
        }

        private void RemoveFile(string fileFolder, string fileName)
        {
            try
            {
                if (IsShortcutOfRebootLoopExists)
                {
                    System.IO.File.Delete(Path.Combine(fileFolder, fileName));
                    Console.WriteLine($"Deleted {fileName} from {fileFolder}");
                }
                else
                {
                    Console.WriteLine($"File {fileName} not found");
                }
            }
            catch (IOException ioExp)
            {
                Console.WriteLine(ioExp.Message);
            }
            Thread.Sleep(1500);
        }

        private void RenameLogFile()
        {
            string oldfileName = $"{AppPath}\\{Constants.LogFileName}";
            string newFileNameTitle = $"{AppPath}\\rebootLog rebooted-{RebootCount}-times {DateTime.Now:MM-dd-yyyy_hh-mm-ss}.txt";
            System.IO.File.Move(oldfileName, newFileNameTitle);
            Thread.Sleep(500);
        }

        private string GetStartupPath()
        {
            string path = Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName;
            if (Environment.OSVersion.Version.Major >= 6)
            {
                path = Directory.GetParent(path).ToString() + @"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup";
            }
            return path;
        }

        public void CreateShortcut(string directoryFullPath, string fileName)
        {
            if (!IsShortcutOfRebootLoopExists)
            {
                WshShell shell = new WshShell();
                string shortcutPathLink = directoryFullPath + $@"\{fileName}.lnk";
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(directoryFullPath + $@"\{fileName}.lnk");
                shortcut.Description = "New shortcut for a RebootLoop";
                shortcut.WorkingDirectory = directoryFullPath;
                shortcut.TargetPath = directoryFullPath + $@"\{Constants.AppName}";
                shortcut.Save();
            }
        }

        #endregion

        private void RunCMDCommand(string CMDCommand)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = CMDCommand;
            process.StartInfo = startInfo;
            process.Start();
        }

        private void SetAutoLogOn()
        {
            RunCMDCommand($"REG ADD {"HKLM\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon"} /v AutoAdminLogon /t REG_SZ /d 1 /f");
            RunCMDCommand($"REG ADD {"HKLM\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon"} /v DefaultDomainName /t REG_SZ /d INTOWINDOWS /f");
            RunCMDCommand($"REG ADD {"HKLM\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon"} /v DefaultUserName /t REG_SZ /d admin /f");
            RunCMDCommand($"REG ADD {"HKLM\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon"} /v DefaultPassword /t REG_SZ /d 1 /f");
        }
    }
}
