using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Plugins.ExcelConvertToJson.Define;
using UnityEngine;

namespace ExcelConvert.Controller
{
    public class BinaryConvertStategy : IConverterStategy
    {
        public void Convert(string fileName, BaseContainer container)
        {
            string path = PathDefine.DataPath + fileName + PostfixDefine.Binary;
            using (FileStream fs = new FileStream(path,FileMode.Create, FileAccess.Write))
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                binaryFormatter.Serialize(fs, container);
                fs.Flush();
                fs.Close();
            }
        }
    }
}