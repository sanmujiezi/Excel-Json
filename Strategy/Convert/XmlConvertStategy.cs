using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using LitJson;
using Plugins.ExcelConvertToJson.Define;
using UnityEngine;

namespace ExcelConvert.Controller
{
    public class XmlConvertStategy : IConverterStategy
    {
        public void Convert(string fileName, BaseContainer container)
        {
            if (!Directory.Exists(PathDefine.DataPath))
            {
                Directory.CreateDirectory(PathDefine.DataPath);
            }
            
            XmlSerializer serializer = new XmlSerializer(typeof(BaseContainer));
            string path = PathDefine.DataPath + fileName + PostfixDefine.Xml;
            using (FileStream fs = new FileStream(path,FileMode.Create,FileAccess.Write))
            {
                StreamWriter writer = new StreamWriter(fs);
                serializer.Serialize(writer,container);                
                
                writer.Flush();
            }
            
            
        }
    }
}
