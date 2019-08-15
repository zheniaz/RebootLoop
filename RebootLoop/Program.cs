using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RebootLoop
{

    // Exercise #1
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("REBOOT SYSTEM, ZHENIA Zgurovets Inc.");
            RebootSystem _rebootSystem = new RebootSystem();
            
            if (_rebootSystem.RebootCount > 10)
            {
                return;
            }
            
            if (_rebootSystem.IsLogFileExists)
            {
                _rebootSystem.Reboot();
            }
            else
            {
                _rebootSystem.CheckIfWantToReboot();
            }
        }
    }

    class Constants
    {
        public const string AppName = "RebootLoop.exe";
        public const string LogFileName = "rebootLog.txt";
        public const string RebootLoopShortcutName = "RebootLoop - Shortcut";
    }
}
