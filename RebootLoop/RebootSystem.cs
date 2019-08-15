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

        public bool IsLogFileExists { get { return CheckIfFileExists($"{AppPath}\\{Constants.LogFileName}"); } set { } } // { get { return CheckIsLogFileExists(); } set { } }
        public bool IsShortcutOfRebootLoopExists { get { return CheckIfFileExists($"{StartupDirectoryFullPath}\\{Constants.RebootLoopShortcutName}.lnk"); } }

        public RebootSystem()
        {
            this.RebootCount = GetRebootCount();
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
                        Console.WriteLine(@"Create Shortcut of RebootLoop, win+r 'shell:startup', and move shortcut to your startup forlder");
                        // CreateShortcut(StartupDirectoryFullPath, Constants.RebootLoopShortcutName);
                        CreateShortCut2(StartupDirectoryFullPath, AppFullPath, Constants.RebootLoopShortcutName);
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
            //CheckYesOrNo("In Reboot(), go on?");
            if (RebootCount > 0 && RebootCount <= NeedToReboot)
            {
                ReadAndRewriteRebootCountInLogFile();
                AddEntryToLogFile();
            }
            if (NeedToReboot == GetRebootsCount())
            {
                FinishRebooting();
                return;
            }
            string time = (NeedToReboot - RebootCount == 1) ? "time" : "times";
            int timesLeftToReboot = 0;
            Console.WriteLine($"{NeedToReboot - RebootCount} {time} left to reboot the system");
            Console.WriteLine("Rebooting System...");
            Thread.Sleep(3000);
            //System.Diagnostics.Process.Start("ShutDown", "-r -t 0");
        }

        public bool CheckIsLogFileExists()
        {
            string filePath = $"{AppPath}\\{Constants.LogFileName}";
            return System.IO.File.Exists(filePath);
        }

        public bool CheckIfFileExists(string filePath)
        {
            Console.WriteLine($"In CheckIfFileExists, path: {filePath}");
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

        public int GetRebootsCount()
        {
            string text = System.IO.File.ReadAllText($"{Constants.LogFileName}");
            string rebootCountString = Regex.Match(text, @"(?<=\|)(.*?)(?=\|)").ToString();
            int rebooted = 0;
            Int32.TryParse(rebootCountString, out rebooted);
            return rebooted;
        }

        public void AddShortcutToSturtup(string path)
        {
            using (PowerShell PowerShell = PowerShell.Create())
            {
                PowerShell.AddScript("$s1 = 'test1'; $s2 = 'test2'; $s1; write-error 'some error';start-sleep -s 7; $s2");

                PSDataCollection<PSObject> outputCollection = new PSDataCollection<PSObject>();
                IAsyncResult result = PowerShell.BeginInvoke();
                PowerShell.AddParameter("param1", "parameter 1 value!");
            }
        }

        private int GetRebootCount()
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

        private void FinishRebooting()
        {
            RenameLogFile();
            // Removing ShortCut from Sturtup folder
            RemoveFile(StartupDirectoryFullPath, Constants.RebootLoopShortcutName + ".lnk");
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
            Thread.Sleep(500);
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

        public void CreateShortCut2(string shorcutPath, string shortcutTarget, string ShortCutName)
        {
            if (!IsShortcutOfRebootLoopExists)
            {
                byte[] bytes = null;

                // Disable impersonation
                using (System.Security.Principal.WindowsImpersonationContext ctx = System.Security.Principal.WindowsIdentity.Impersonate(IntPtr.Zero))
                {
                    // Get a temp file name (the shell commands won't work without .lnk extension)
                    var path = Path.GetTempPath();
                    string temp = Path.Combine(shorcutPath, ShortCutName + ".lnk");
                    try
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(temp);
                        shortcut.TargetPath = shortcutTarget;
                        shortcut.Save();
                        bytes = System.IO.File.ReadAllBytes(temp);
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine(exception.Message);
                    }
                }
            }
        }

        #region This way not working, need to investigate and fix

        public bool CreateShortcut(string directoryFullPath, string fileName)
        {
            if (!IsShortcutOfRebootLoopExists)
            {
                object shDesktop = (object)directoryFullPath;
                WshShell shell = new WshShell();
                string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + $@"\{fileName}.lnk";
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
                shortcut.Description = "New shortcut for a RebootLoop";
                shortcut.Hotkey = "Ctrl+Shift+N";
                shortcut.TargetPath = directoryFullPath + $@"\{Constants.AppName}";

                AddDirectorySecurity(directoryFullPath, Environment.UserDomainName + "\\" + Environment.UserName, FileSystemRights.FullControl, AccessControlType.Allow);
                //SetAccessRule(directoryFullPath);

                shortcut.Save();
            }

            return IsShortcutOfRebootLoopExists;
        }

        private void AddDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            // Create a new DirectoryInfo object.
            DirectoryInfo dInfo = new DirectoryInfo(FileName);

            // Get a DirectorySecurity object that represents the             current security settings.
            DirectorySecurity dSecurity = dInfo.GetAccessControl();

            // Add the FileSystemAccessRule to the security settings.
           dSecurity.AddAccessRule(new FileSystemAccessRule(Account, Rights, ControlType));

            // Set the new access settings.
           dInfo.SetAccessControl(dSecurity);
        }

        private void SetAccessRule(string directory)
        {
            DirectorySecurity sec = Directory.GetAccessControl(directory);
            FileSystemAccessRule accRule = new FileSystemAccessRule(Environment.UserDomainName + "\\" + Environment.UserName, FileSystemRights.FullControl, AccessControlType.Allow);
            sec.AddAccessRule(accRule);
        }

        #endregion
    }
}
