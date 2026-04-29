using ACSMCOMPONENTS24Lib;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class ComHelper
    {
        public static T GetObject<T>(IAcSmObjectId id) where T : class
        {
            return TryHelper.Run(() =>
            {
                if (id == null)
                {
                    Log.Warn("GetObject<T> kreeg een null objectId.");
                    return null;
                }

                var obj = id.GetPersistObject();
                return obj as T;
            }, "ComHelper.GetObject");
        }
    }
}

    
