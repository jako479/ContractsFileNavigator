using System;
using System.IO;
using System.Windows.Forms;

namespace ContractsFileNavigator
{
    /// <summary>
    /// Appends every caught error and any unhandled exception to %TEMP%\ContractsFileNavigator.log.
    /// Never throws; logging must not be able to break the add-in.
    /// </summary>
    internal static class Diagnostics
    {
        /// <summary>Once the log passes this size, only its newest <see cref="KeepBytes"/> are kept.</summary>
        private const long MaxLogBytes = 1024 * 1024;
        private const int KeepBytes = 256 * 1024;

        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "ContractsFileNavigator.log");
        private static readonly object Gate = new object();

        /// <summary>The last message written; an identical one straight after it is dropped.</summary>
        private static string lastMessage;

        /// <summary>
        /// Appends one timestamped line. A message identical to the previous one is dropped, so a
        /// failure that repeats on every timer tick fills one line, not the file. This is the one place
        /// a failure stays silent: there is nowhere left to report a logging failure.
        /// </summary>
        public static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    if (string.Equals(message, lastMessage, StringComparison.Ordinal)) return;
                    lastMessage = message;

                    TrimIfLarge();
                    File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}{Environment.NewLine}");
                }
            }
            catch { }
        }

        /// <summary>
        /// Records exceptions that nothing else caught, so a crash leaves its name and stack behind.
        /// </summary>
        public static void HookUnhandledExceptions()
        {
            try
            {
                AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
                    Write("UNHANDLED (AppDomain): " + e.ExceptionObject);

                Application.ThreadException += (sender, e) =>
                    Write("UNHANDLED (WinForms): " + e.Exception);
            }
            catch { }
        }

        /// <summary>
        /// Keeps a repeating error from growing the file without limit.
        /// </summary>
        private static void TrimIfLarge()
        {
            FileInfo info = new FileInfo(LogPath);
            if (!info.Exists || info.Length <= MaxLogBytes) return;

            byte[] all = File.ReadAllBytes(LogPath);
            byte[] tail = new byte[KeepBytes];
            Array.Copy(all, all.Length - KeepBytes, tail, 0, KeepBytes);
            File.WriteAllBytes(LogPath, tail);
        }
    }
}
