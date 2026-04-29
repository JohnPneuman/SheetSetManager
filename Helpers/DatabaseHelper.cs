using ACSMCOMPONENTS24Lib;
using System;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class DatabaseHelper
    {
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
                    Log.Debug("[SSM] Database gelockt.");
                    return true;
                }

                Log.Warn("[SSM] Database was al gelockt.");
                return true;
            }, "DatabaseHelper.TryLock", false);
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
