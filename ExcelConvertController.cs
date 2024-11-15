using System;
using System.Collections.Generic;
using System.IO;
using LitJson;
using OfficeOpenXml;
using UnityEditor;
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
            public ConvertType convertType { get; private set; }
            private IReadStrategy readStrategy;
            private IConverterStategy converterStategy;

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

            public void CreateModel()
            {
                SelectStrategy();
                if (converterStategy == null || convertType == ConvertType.None)
                {
                    return;
                }

                //var tempVlaue= excelPath.Replace("/", "\\");
                string[] xlsxFiles = Directory.GetFiles(excelPath, "*.xlsx");
                
                
                foreach (var value in xlsxFiles)
                {
                    readStrategy.CreateModel(value);
                }
            }

            public void ConvertData()
            {
                int strategyInt = PlayerPrefs.GetInt("Strategy");
                convertType = (ConvertType)strategyInt;

                SelectStrategy();

                if (converterStategy == null)
                {
                    return;
                }

                excelPath = PlayerPrefs.GetString("excelPath");

                string[] xlsxFiles = Directory.GetFiles(excelPath, "*.xlsx");

                foreach (var value in xlsxFiles)
                {
                    var allTabls = readStrategy.ReadExcel(value);
                    Debug.Log($"读取成功,共{allTabls.Count}个表");
                    foreach (var VARIABLE in allTabls)
                    {
                        Debug.Log($"表名为：{VARIABLE.Key}");
                        converterStategy.Convert(VARIABLE.Key, VARIABLE.Value);
                    }
                }
                AssetDatabase.Refresh();
                
            }


            private void SelectStrategy()
            {
                readStrategy = new ExcelReadStrategy();

                switch (convertType)
                {
                    case ConvertType.None:
                        Debug.Log("没有选择传唤类型");
                        return;
                    case ConvertType.Binary:
                        converterStategy = new BinaryConvertStategy();
                        break;
                    case ConvertType.Json:
                        converterStategy = new JsonConvertStategy();
                        break;
                    case ConvertType.Xml:
                        converterStategy = new XmlConvertStategy();
                        return;
                }
            }

            public void SetExcelPath(string path)
            {
                excelPath = path;
            }


            public void SetConvertType(ConvertType convertType)
            {
                this.convertType = convertType;
            }

            public void SaveSeletionData()
            {
                if (!string.IsNullOrEmpty(excelPath))
                {
                    PlayerPrefs.SetString("excelPath", excelPath);
                }

                if (convertType != null || convertType != ConvertType.None)
                {
                    PlayerPrefs.SetInt("Strategy", (int)convertType);
                }
            }

            public void LoadSelectionData()
            {
                if (PlayerPrefs.HasKey("excelPath"))
                {
                    excelPath = PlayerPrefs.GetString("excelPath");
                }

                if (PlayerPrefs.HasKey("Strategy"))
                {
                    int strategyInt = PlayerPrefs.GetInt("Strategy");
                    convertType = (ConvertType)strategyInt;
                }
            }

            public void CreatePluginEnvironment()
            {
                //检查Scripts目录下是否存在
            }
        }
    }
}