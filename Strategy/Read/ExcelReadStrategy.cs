using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel;
using Plugins.ExcelConvertToJson.Tools;
using UnityEngine;

namespace ExcelConvert.Controller
{
    public class ExcelReadStrategy : IReadStrategy
    {
        public void ReadData(string path, out string modelName, out string containerModelName)
        {
            modelName = "";
            containerModelName = "";
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                DataSet result = excelReader.AsDataSet();
                if (result.Tables.Count < 0)
                {
                    Debug.Log($"Excel文件为空");
                    return;
                }

                for (int i = 0; i < result.Tables.Count; i++)
                {
                    string name = result.Tables[i].TableName;
                    if (!ExcelHandel.IsExistClass(name))
                    {
                        ExcelHandel.CreateModel(result.Tables[i],name);
                    }    
                }
                

                //读取文件中所有的表
                for (int i = 0; i < result.Tables.Count; i++)
                {
                    DataTable table = result.Tables[i];
                    string name = table.TableName;
                    
                    if (!ExcelHandel.IsExistClass(name))
                    {
                        ExcelHandel.CreateModel(table,name);
                    }
                    
                    Type type = Type.GetType(name);
                    var dic = ExcelHandel.CreateContainer(table,type);
                }

                Debug.Log($"从{path}读取了Excel文件");
            }
        }
    }
}