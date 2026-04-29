using System;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class TryHelper
    {
        public static void Run(Action action, string context = "")
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                Log.Error($"[Fout] {context}: {ex.Message}\n{ex.StackTrace}");
            }
        }

        public static T? Run<T>(Func<T?> func, string context = "", T? fallback = default)
        {
            try
            {
                return func();
            }
            catch (Exception ex)
            {
                Log.Error($"[Fout] {context}: {ex.Message}\n{ex.StackTrace}");
                return fallback;
            }
        }
    }
}
