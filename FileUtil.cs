using Newtonsoft.Json;
using System.IO;

namespace GoogleExcel
{
    internal static class FileUtil
    {
        public static T JsonFile<T>(string fileName)
        {
            
            using (StreamReader r = new StreamReader(fileName))
            {
                string json = r.ReadToEnd();
                return JsonConvert.DeserializeObject<T>(json);
            }
        }
    }
}
