using System;
using System.IO;

namespace AESCConstruct2026.FrameGenerator.Utilities
{
    /// <summary>Severity level for log messages.</summary>
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    /// <summary>File-based diagnostic logger that appends timestamped messages to the shared log file.</summary>
    public static class Logger
    {
        private static readonly string appDataDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AESCConstruct");

        private static readonly string logPath =
            Path.Combine(appDataDir, "AESCConstruct2026_Log.txt");

        private static readonly object _lock = new object();

        private const long MaxLogSize = 10 * 1024 * 1024; // 10 MB

        /// <summary>Minimum severity level to write. Messages below this level are discarded.</summary>
        public static LogLevel MinLevel { get; set; } = LogLevel.Debug;

        /// <summary>Appends a timestamped log entry to the shared diagnostic log file.</summary>
        public static void Log(string message)
        {
            Log(LogLevel.Info, message);
        }

        /// <summary>Appends a timestamped log entry with the specified severity level.</summary>
        public static void Log(LogLevel level, string message)
        {
            if (level < MinLevel)
                return;

            try
            {
                string line;
                lock (_lock)
                {
                    line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [" + level + "] " + message + "\n";
                }

                // B4: Ensure directory exists before writing
                Directory.CreateDirectory(appDataDir);

                // B12: Rotate if file exceeds 10 MB
                RotateIfNeeded();

                File.AppendAllText(logPath, line);
            }
            catch
            {
                // Logger should never crash the host application
            }
        }

        private static void RotateIfNeeded()
        {
            try
            {
                var info = new FileInfo(logPath);
                if (info.Exists && info.Length > MaxLogSize)
                {
                    string backupPath = logPath + ".bak";
                    File.Copy(logPath, backupPath, overwrite: true);
                    File.Delete(logPath);
                }
            }
            catch
            {
                // Best-effort rotation; don't let it break logging
            }
        }
    }
}
