using log4net;
using BoekSolutions.SheetSetEditor.Helpers;
using BoekSolutions.SheetSetEditor;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class Log
    {
        private static readonly ILog log = LogManager.GetLogger("SheetSetLogger");

        private static void WriteToEditor(string prefix, string message)
        {
            TryHelper.Run(() =>
            {
                Aliases.Ed?.WriteMessage($"\n{prefix} {message}");
            }, $"Log naar editor [{prefix}]");
        }

        public static void Info(string message)
        {
            TryHelper.Run(() =>
            {
                if (log.IsInfoEnabled) log.Info(message);
                WriteToEditor("[INFO]", message);
            }, "Log.Info");
        }

        public static void Debug(string message)
        {
            TryHelper.Run(() =>
            {
                if (log.IsDebugEnabled) log.Debug(message);
                WriteToEditor("[DEBUG]", message);
            }, "Log.Debug");
        }

        public static void Error(string message)
        {
            TryHelper.Run(() =>
            {
                if (log.IsErrorEnabled) log.Error(message);
                WriteToEditor("[ERROR]", message);
            }, "Log.Error");
        }

        public static void Warn(string message)
        {
            TryHelper.Run(() =>
            {
                if (log.IsWarnEnabled) log.Warn(message);
                WriteToEditor("[WARN]", message);
            }, "Log.Warn");
        }
    }
}
