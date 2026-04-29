using System;
using System.IO;
using System.Reflection;
using log4net;
using log4net.Config;

/// <summary>
/// Init klasse voor log4net logging. Wordt één keer aangeroepen bij het starten van de plugin.
/// Let op: AutoCAD is de host, dus AppDomain.BaseDirectory wijst naar de AutoCAD-installatie.
/// Daarom gebruiken we Assembly.GetExecutingAssembly() om het pad van onze eigen plugin te vinden.
/// </summary>
public static class LoggerInitializer
{
    private static bool _isInitialized = false;

    public static void Init()
    {
        if (_isInitialized) return;

        try
        {
            // 👇 Gebruik de locatie van je .dll, NIET AppDomain.CurrentDomain.BaseDirectory (die wijst naar acad.exe!)
            string exeFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
            string configPath = Path.Combine(exeFolder, "Config", "log4net.config");

            System.Diagnostics.Debug.WriteLine("👀 Verwachte config path: " + configPath);
            System.Diagnostics.Debug.WriteLine("📦 Bestaat config? " + File.Exists(configPath));

            if (!File.Exists(configPath))
            {
                System.Diagnostics.Debug.WriteLine("💥 log4net.config NIET GEVONDEN: " + configPath);
                return;
            }

            // 📂 Zorg dat logbestanden in AppData worden geplaatst
            string logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BoekSolutions", "Logs"
            );
            Directory.CreateDirectory(logDir);
            string logFile = Path.Combine(logDir, "ssm.log");

            // 🔗 Zet logpad property (werkt samen met PatternString in log4net.config)
            GlobalContext.Properties["LogFilePath"] = logFile;

            // ✅ Configureer log4net voor deze assembly, zodat hij niet afhankelijk is van acad.exe.config
            var repoThisAssembly = LogManager.GetRepository(Assembly.GetExecutingAssembly());
            XmlConfigurator.Configure(repoThisAssembly, new FileInfo(configPath));

            System.Diagnostics.Debug.WriteLine("✅ log4net is succesvol geconfigureerd.");
            System.Diagnostics.Debug.WriteLine("📁 Logbestand: " + logFile);
            System.Diagnostics.Debug.WriteLine("⚙️ Config uit: " + configPath);

            _isInitialized = true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("❌ Fout bij log-init: " + ex.Message);
        }
    }
}
