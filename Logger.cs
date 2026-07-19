using System;
using System.IO;

namespace ExcelSplitter
{
    /// <summary>
    /// Writes operational messages to both the console (in color) and a log file.
    /// Each run creates a new timestamped log file.
    /// </summary>
    public class Logger : IDisposable
    {
        private readonly StreamWriter _writer;
        private readonly object _lock = new();
        private bool _disposed;

        public Logger(string logFolder)
        {
            Directory.CreateDirectory(logFolder);
            string logFileName = $"log_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            LogFilePath = Path.Combine(logFolder, logFileName);
            _writer = new StreamWriter(LogFilePath, append: true) { AutoFlush = true };
        }

        public string LogFilePath { get; }

        public void Info(string message) => Write("INFO", message, ConsoleColor.Gray);
        public void Warning(string message) => Write("WARN", message, ConsoleColor.Yellow);
        public void Error(string message) => Write("ERROR", message, ConsoleColor.Red);
        public void Success(string message) => Write("OK", message, ConsoleColor.Green);

        private void Write(string level, string message, ConsoleColor color)
        {
            lock (_lock)
            {
                if (_disposed) return;

                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";

                var prevColor = Console.ForegroundColor;
                Console.ForegroundColor = color;
                Console.WriteLine(line);
                Console.ForegroundColor = prevColor;

                _writer.WriteLine(line);
            }
        }

        public void Dispose()
        {
            lock (_lock)
            {
                if (_disposed) return;
                _disposed = true;
                _writer.Dispose();
            }
        }
    }
}
