using System;
using System.IO;

namespace AESCConstruct25.Utilities
{
    public static class Logger
    {
        private static readonly string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AESCConstruct25_Log.txt");

        public static void Log(string message)
        {
            File.AppendAllText(logPath, DateTime.Now + ": " + message + "\n");
        }
    }
}
