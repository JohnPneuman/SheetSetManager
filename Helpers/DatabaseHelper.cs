using ACSMCOMPONENTS24Lib;
using System;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class DatabaseHelper
    {
        /// <summary>
        /// Probeert de database te locken. Geeft terug of WIJ de lock gezet hebben.
        /// Als de database al gelockt was door AutoCAD, returnt dit false (niet door ons gelockt).
        /// </summary>
        public static bool TryLock(IAcSmDatabase db)
        {
            return TryHelper.Run(() =>
            {
                if (db == null)
                {
                    Log.Error("[SSM] Kan database niet locken: db is null.");
                    return false;
                }

                if (db.GetLockStatus() != AcSmLockStatus.AcSmLockStatus_Locked_Local)
                {
                    db.LockDb(db);
                    Log.Debug("[SSM] Database gelockt door plugin.");
                    return true;
                }

                Log.Warn("[SSM] Database was al gelockt (door AutoCAD). Plugin neemt lock NIET over.");
                return false;
            }, "DatabaseHelper.TryLock", false);
        }

        /// <summary>
        /// Slaat de database op zonder te unlocken. Gebruik dit als AutoCAD de lock bezit.
        /// </summary>
        public static void SaveWithoutUnlock(IAcSmDatabase db)
        {
            TryHelper.Run(() =>
            {
                if (db == null)
                {
                    Log.Error("[SSM] SaveWithoutUnlock afgebroken: db is null.");
                    return;
                }

                SaveHelper.SaveDatabase(db);

                db.Save(null);
                Log.Info("[SSM] Database opgeslagen (zonder unlock).");
            }, "DatabaseHelper.SaveWithoutUnlock");
        }

        public static void UnlockDb(IAcSmDatabase db, bool saveBeforeUnlock = false)
        {
            TryHelper.Run(() =>
            {
                if (db == null)
                {
                    Log.Info("[SSM] Kan database niet unlocken: db is null.");
                    return;
                }

                if (saveBeforeUnlock)
                {
                    Log.Info("[SSM] Save vóór unlock (via UnlockDb)...");

                    SaveHelper.SaveDatabase(db);
                }

                db.UnlockDb(db, true);
                Log.Info("[SSM] Database ge-unlocked.");
            });
        }

        public static bool IsLocked(IAcSmDatabase db)
        {
            var result = false;

            TryHelper.Run(() =>
            {
                if (db == null)
                {
                    Log.Info("[SSM] IsLocked: db is null.");
                    return;
                }

                var status = db.GetLockStatus();
                Log.Info("[SSM] Lock status = " + status.ToString());

                result = (status == AcSmLockStatus.AcSmLockStatus_Locked_Local);
            });

            return result;
        }
    }
}
