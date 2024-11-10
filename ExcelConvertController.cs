using System;
using System.Collections.Generic;
using System.IO;
using LitJson;
using OfficeOpenXml;
using UnityEngine;

namespace ExcelConvert
{
    public enum ConvertType
    {
        None = 0,
        Json = 1,
        Binary = 2,
        Xml = 3
    }

    namespace Controller
    {
        public class ExcelConvertController
        {
            public string excelPath { get; private set; }
            public string outputPath { get; private set; }
            public ConvertType convertType { get; private set; }

            private static ExcelConvertController instance;

            public static ExcelConvertController Instance
            {
                get
                {
                    if (instance == null)
                    {
                        instance = new ExcelConvertController();
                    }
                    
                    return instance;
                }
            }

            public ExcelConvertController()
            {
                
            }

            #region 旧的Excel转换Json方案

            public void ConvertToJson()
            {
                if (excelPath == null || outputPath == null || excelPath == "" || outputPath == "")
                {
                    Debug.LogWarning("》》请选择 输入文件 或者 输出地址 !");
                    return;
                }

                if (!File.Exists(excelPath))
                {
                    return;
                }


                FileInfo fileInfo = new FileInfo(this.excelPath);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    int rows = worksheet.Dimension.Rows;
                    int columns = worksheet.Dimension.Columns;
                    //判断接收实例
                    List<Dictionary<string, object>> dataLists = new List<Dictionary<string, object>>();
                    for (int i = 2; i <= rows; i++)
                    {
                        Dictionary<string, object> cell = new Dictionary<string, object>();
                        for (int j = 1; j <= columns; j++)
                        {
                            string columnName = worksheet.Cells[1, j].Value.ToString();
                            object cellValue = GetCellType(worksheet.Cells[i, j].Value.ToString(), columnName);
                            cell.Add(columnName, cellValue);
                        }

                        dataLists.Add(cell);
                    }

                    //生成json
                    if (dataLists != null)
                    {
                        string json = JsonMapper.ToJson(dataLists);
                        File.WriteAllText(outputPath + "/" + worksheet.Name + ".json", json);
                        Debug.Log("写入成功" + outputPath + "/" + worksheet.Name + ".json");
                    }
                }
            }

            public object GetCellType(string cell, string columnName)
            {
                if (excelPath == null || excelPath == "" || columnName == null || columnName == "")
                {
                    return "";
                }

                switch (columnName)
                {
                    case "id":
                        return int.Parse(cell);
                    case "name":
                    case "image":
                    case "descripts":
                        return cell;
                    default:
                        return cell;
                }

                return null;
            }

            #endregion

            public void Convert()
            {
                IReadStrategy readStrategy = new ExcelReadStrategy();
                ConvertReader reader = new ConvertReader(readStrategy);
                
                string modelName = "";
                string containerModelName = "";
                reader.ReadData(excelPath,out modelName,out containerModelName);
                
                
                
                
            }
            

            public void SetExcelPath(string path)
            {
                excelPath = path;
            }

            public void SetOutputPath(string path)
            {
                outputPath = path;
            }

            public void SetConvertType(ConvertType convertType)
            {
                this.convertType = convertType;
            }
        }
    }
}