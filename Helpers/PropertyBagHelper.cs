using System;
using System.Collections.ObjectModel;
using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.ViewModels;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class PropertyBagHelper
    {
        public static ObservableCollection<CustomPropertyViewModel> GetCustomProperties(IAcSmObjectId id)
        {
            var result = new ObservableCollection<CustomPropertyViewModel>();

            TryHelper.Run(() =>
            {
                using (var scope = new ComScope())
                {
                    var persist = id.GetPersistObject();

                    IAcSmCustomPropertyBag bag = null;

                    if (persist is IAcSmSheet sheet)
                    {
                        scope.Track(sheet);
                        bag = scope.Track(sheet.GetCustomPropertyBag());
                        Log.Debug("[PBH] persist is IAcSmSheet → propertybag opgehaald.");
                    }
                    else if (persist is IAcSmCustomPropertyBag directBag)
                    {
                        bag = scope.Track(directBag);
                        Log.Debug("[PBH] persist is IAcSmCustomPropertyBag → direct gebruikt.");
                    }
                    else
                    {
                        Log.Warn($"[PBH] persist is van type {persist.GetType().Name} → geen propertybag.");
                    }

                    if (bag != null)
                    {
                        var enumProps = scope.Track(bag.GetPropertyEnumerator());
                        string propName;
                        AcSmCustomPropertyValue temp;
                        int count = 0;

                        while (true)
                        {
                            enumProps.Next(out propName, out temp);
                            if (string.IsNullOrEmpty(propName)) break;

                            scope.Track(temp);
                            var value = temp?.GetValue() as string;

                            result.Add(new CustomPropertyViewModel
                            {
                                PropertyName = propName,
                                PropertyValue = value ?? ""
                            });

                            count++;
                        }

                        Log.Info($"[PBH] {count} custom properties geladen.");
                    }
                    else
                    {
                        Log.Warn("[PBH] Custom property bag is null.");
                    }
                }
            });

            return result;
        }

        public static string GetCustomProperty(IAcSmObjectId id, string key)
        {
            var props = GetCustomProperties(id);
            foreach (var prop in props)
            {
                if (string.Equals(prop.PropertyName, key, StringComparison.OrdinalIgnoreCase))
                    return prop.PropertyValue;
            }
            return null;
        }

        public static void SetCustomPropertyIfValid(IAcSmObjectId id, string name, string value)
        {
            TryHelper.Run(() =>
            {
                if (string.IsNullOrEmpty(name)) return;

                using (var scope = new ComScope())
                {
                    var persist = id.GetPersistObject();

                    IAcSmCustomPropertyBag bag = null;

                    if (persist is IAcSmSheet sheet)
                    {
                        scope.Track(sheet);
                        bag = scope.Track(sheet.GetCustomPropertyBag());
                    }
                    else if (persist is IAcSmCustomPropertyBag directBag)
                    {
                        bag = scope.Track(directBag);
                    }

                    if (bag == null)
                    {
                        Log.Warn($"[PBH] PropertyBag is null → property '{name}' wordt niet geschreven.");
                        return;
                    }

                    var enumProps = scope.Track(bag.GetPropertyEnumerator());
                    enumProps.Reset();

                    string currentName;
                    AcSmCustomPropertyValue currentValue;

                    bool found = false;

                    while (true)
                    {
                        enumProps.Next(out currentName, out currentValue);
                        if (string.IsNullOrEmpty(currentName)) break;

                        scope.Track(currentValue);

                        if (string.Equals(currentName, name, StringComparison.OrdinalIgnoreCase))
                        {
                            if (currentValue is IAcSmCustomPropertyValue editableValue)
                            {
                                editableValue.SetValue(value);
                                Log.Debug($"[PBH] Bestaande property overschreven: {name} = {value}");
                                found = true;
                                break;
                            }
                        }
                    }

                    // Als niet gevonden: niets doen (of hier kun je een "create new" poging doen als je dat toch wil)
                    if (!found)
                    {
                        Log.Warn($"[PBH] Property '{name}' niet gevonden in bag → niet overschreven.");
                    }
                }
            });
        }

    }
}