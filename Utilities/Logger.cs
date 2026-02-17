using System;
using System.IO;

namespace AESCConstruct2026.FrameGenerator.Utilities
{
    /// <summary>File-based diagnostic logger that appends timestamped messages to the shared log file.</summary>
    public static class Logger
    {
        private static readonly string appDataDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AESCConstruct");

        private static readonly string logPath =
            Path.Combine(appDataDir, "AESCConstruct2026_Log.txt");

        private static readonly object _lock = new object();

        /// <summary>Appends a timestamped log entry to the shared diagnostic log file.</summary>
        public static void Log(string message)
        {
            lock (_lock)
            {
                File.AppendAllText(logPath, DateTime.Now + ": " + message + "\n");
            }
        }
    }
}
