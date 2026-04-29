using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Autodesk.AutoCAD.Runtime;
using static BoekSolutions.SheetSetEditor.Aliases;
using BoekSolutions.SheetSetEditor.Helpers;

[assembly: ExtensionApplication(typeof(BoekSolutions.SheetSetEditor.BoekSheetSetPlugin))]

namespace BoekSolutions.SheetSetEditor
{
    public class BoekSheetSetPlugin : IExtensionApplication
    {
        public void Initialize()
        {
            LoggerInitializer.Init();
            Log.Info("[SSM] Plugin geïnitialiseerd (.NET 10, AutoCAD 2027).");
        }

        public void Terminate()
        {
            Log.Info("[SSM] Plugin afgesloten.");
        }
    }

    public class Commands
    {
        [CommandMethod("SSM_UI")]
        public static void StartUI()
        {
            Log.Info("[SSM] Start standalone SheetSet Editor...");

            string pluginDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
            string appExe = Path.Combine(pluginDir, "SheetSetEditor.App.exe");
            if (!File.Exists(appExe))
            {
                Log.Error($"[SSM] SheetSetEditor.App.exe niet gevonden: {appExe}");
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo { FileName = appExe, UseShellExecute = false });
                Log.Info($"[SSM] App gestart: {appExe}");
            }
            catch (System.Exception ex)
            {
                Log.Error($"[SSM] Kon App niet starten: {ex.Message}");
            }
        }
    }
}
