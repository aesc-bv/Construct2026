using System;
using System.IO;

namespace AESCConstruct25.FrameGenerator.Utilities
{
    public static class Logger
    {
        //private static readonly string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AESCConstruct25_Log.txt");
        private static readonly string appDataDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AESCConstruct");

        private static readonly string logPath =
            Path.Combine(appDataDir, "AESCConstruct25_Log.txt");

        public static void Log(string message)
        {
            File.AppendAllText(logPath, DateTime.Now + ": " + message + "\n");
        }
    }
}
