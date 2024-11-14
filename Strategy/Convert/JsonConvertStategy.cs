using System.Collections.Generic;
using System.IO;
using LitJson;
using Plugins.ExcelConvertToJson.Define;

namespace ExcelConvert.Controller
{
    public class JsonConvertStategy : IConverterStategy
    {
        public void Convert(string fileName, BaseContainer container)
        {
            if (!Directory.Exists(PathDefine.DataPath))
            {
                Directory.CreateDirectory(PathDefine.DataPath);
            }
            string path = PathDefine.DataPath + fileName + PostfixDefine.Json;
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                StreamWriter sw = new StreamWriter(fs);
                string toJson = JsonMapper.ToJson(container);
                sw.Write(toJson);
                sw.Flush();
                fs.Close();
            }
            //
            // FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            // StreamWriter sw = new StreamWriter(fs);
            // string toJson = JsonMapper.ToJson(container);
            // sw.Write(toJson);
            // sw.Flush();
            // fs.Flush();
            // fs.Close();
        }
    }
}