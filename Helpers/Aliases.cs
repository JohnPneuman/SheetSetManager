using System;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcDoc = Autodesk.AutoCAD.ApplicationServices.Document;
using AcEd = Autodesk.AutoCAD.EditorInput;

namespace BoekSolutions.SheetSetEditor
{
    public static class Aliases
    {
        public static AcDoc? ActiveDoc => AcApp.DocumentManager.MdiActiveDocument;
        public static AcEd.Editor? Ed => ActiveDoc?.Editor;
        public static string AppDataPath => Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
    }
}
