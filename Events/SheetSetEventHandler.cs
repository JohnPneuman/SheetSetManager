using System;
using ACSMCOMPONENTS24Lib;
using Autodesk.AutoCAD.Runtime;
using static BoekSolutions.SheetSetEditor.Aliases;
using BoekSolutions.SheetSetEditor.Helpers;

namespace BoekSolutions.SheetSetEditor.Events
{
    public class SheetSetEventHandler : IAcSmEvents
    {
        void IAcSmEvents.OnChanged(AcSmEvent ev, IAcSmPersist comp)
        {
            TryHelper.Run(() =>
            {
                if (comp == null) return;

                using (var dbScope = new ComScope())
                {
                    var db = dbScope.Track(comp.GetDatabase());
                    string dbName = TryHelper.Run(() => db?.GetFileName()) ?? "<unknown>";

                    switch (ev)
                    {
                        case AcSmEvent.ACSM_DATABASE_OPENED:
                            Log.Info($"[SSM] Opened: {dbName}");
                            break;

                        case AcSmEvent.ACSM_DATABASE_CHANGED:
                            Log.Info($"[SSM] Changed: {dbName}");
                            break;

                        case AcSmEvent.SHEET_DELETED:
                            using (var sheetScope = new ComScope())
                            {
                                var sheet = sheetScope.Track(comp as IAcSmSheet);
                                Log.Info($"[SSM] Sheet deleted: {sheet?.GetName() ?? "<null>"}");
                            }
                            break;

                        case AcSmEvent.SHEET_SUBSET_CREATED:
                            using (var subsetScope = new ComScope())
                            {
                                var subset = subsetScope.Track(comp as IAcSmSubset);
                                Log.Info($"[SSM] Subset created: {subset?.GetName() ?? "<null>"}");
                            }
                            break;

                        case AcSmEvent.SHEET_SUBSET_DELETED:
                            using (var subsetScope = new ComScope())
                            {
                                var subset = subsetScope.Track(comp as IAcSmSubset);
                                Log.Info($"[SSM] Subset deleted: {subset?.GetName() ?? "<null>"}");
                            }
                            break;

                        default:
                            Log.Info($"[SSM] Event: {ev} on {dbName}");
                            break;
                    }
                }

            });
        }
    }
}