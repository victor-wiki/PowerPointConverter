using Newtonsoft.Json;
using System.Reflection;

namespace PowerPointConverter.Helper
{
    public class ObjectHelper
    {
        public static T CloneObject<T>(object obj)
        {
            return (T)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(obj), typeof(T));
        }

        public static void CopyProperties(object source, object target, List<string> excludePropertyNames = null)
        {
            var sourceProps = source.GetType().GetProperties().Where(x => x.CanRead).ToList();
            var targetProps = target.GetType().GetProperties().Where(x => x.CanWrite).ToList();

            foreach (var sourceProp in sourceProps)
            {
                if (excludePropertyNames != null && excludePropertyNames.Contains(sourceProp.Name))
                {
                    continue;
                }

                if (targetProps.Any(x => x.Name == sourceProp.Name))
                {
                    var p = targetProps.FirstOrDefault(x => x.Name == sourceProp.Name);
                    if (p != null && p.CanWrite)
                    {
                        p.SetValue(target, sourceProp.GetValue(source, null), null);
                    }
                }
            }
        }

        public static bool HasProperty(object obj, string propertyName)
        {
            if (obj == null)
            {
                return false;
            }

            var properties = obj.GetType().GetProperties().Where(x => x.CanRead);

            return properties.Any(item => item.Name == propertyName);
        }

        public static bool AreObjectsEqual(object object1, object object2)
        {
            if (object1 == null || object2 == null)
            {
                return object1 == object2;
            }

            Type type1 = object1.GetType();
            Type type2 = object2.GetType();

            PropertyInfo[] properties1 = type1.GetProperties();
            PropertyInfo[] properties2 = type1.GetProperties();

            if (properties1.Length != properties2.Length)
            {
                return false;
            }

            foreach (PropertyInfo property in properties1)
            {
                if (!Equals(property.GetValue(object1), property.GetValue(object2)))
                {
                    return false;
                }
            }

            return true;
        }

        public static object GetValue(object obj, string propertyName)
        {
            if (obj == null || propertyName == null)
            {
                return null;
            }

            var property = obj.GetType().GetProperties().FirstOrDefault(item => item.Name == propertyName);

            if (property != null)
            {
                return property.GetValue(obj);
            }

            return null;
        }

        public static string GetObjectJson(object obj)
        {
            return JsonConvert.SerializeObject(obj);
        }

        public static T GetObjectFromJson<T>(string json)
        {
            if(json == null)
            {
                return default(T);
            }

            return (T)JsonConvert.DeserializeObject<T>(json);
        }
    }
}
