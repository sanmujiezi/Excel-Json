using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel;
using Plugins.ExcelConvertToJson.Tools;
using UnityEditor;
using UnityEngine;

namespace ExcelConvert.Controller
{
    public class ExcelReadStrategy : IReadStrategy
    {
        public void CreateModel(string path)
        {
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
                    DataTable table = result.Tables[i];
                    string name = table.TableName;
                    
                    if (!ExcelHandel.IsExistClass(name))
                    {
                        ExcelHandel.CreateModel(table,name);
                        //ExcelHandel.CreateContainer(name);
                    }
                    
                    AssetDatabase.Refresh();
               
                }
                
                Debug.Log($"从{path}读取到了Excel文件");
            }
        }

        public Dictionary<string,BaseContainer> ReadData(string path)
        {
            Dictionary<string,BaseContainer> fileTable = new Dictionary<string,BaseContainer>();
            
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                DataSet result = excelReader.AsDataSet();
                
                for (int i = 0; i < result.Tables.Count; i++)
                {
                    DataTable table = result.Tables[i];
                    var tableC = ExcelHandel.ReadExcel(table);
                    fileTable.Add(table.TableName,tableC);
                }
            }

            return fileTable;

        }
    }
}