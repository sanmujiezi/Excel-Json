using System.Collections.Generic;
using System.IO;
using ExcelConvert;
using ExcelConvert.Controller;
using Plugins.ExcelConvertToJson.Define;
using UnityEditor;
using UnityEngine;

namespace Plugins.ExcelConvertToJson.Tools
{
    public class ReaderHandel
    {
        
        
        private static ReaderHandel _instance;

        public static ReaderHandel Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ReaderHandel();
                }

                return _instance;
            }
        }
        private IAnalysis anialysis;
        
        public KeyValuePair<string, BaseContainer> GetKeyAndValue<T>(DataFileType type)
        {
            string tableName = typeof(T).Name;
            string postfix = "";
            switch (type)
            {
                case DataFileType.Json:
                    postfix = PostfixDefine.Json;
                    anialysis = new AnalysisJson();                  
                    break;
                case DataFileType.Xml:
                    postfix = PostfixDefine.Xml;
                    anialysis = new AnalysisXml();                    
                    break;
                case DataFileType.Binary:
                    postfix = PostfixDefine.Binary;
                    anialysis = new AnalysisBinary();
                    break;
            }
            
            string dataPath = Application.dataPath + PathDefine.DataPath + tableName + postfix;
            if (!File.Exists(dataPath) )
            {
                Debug.LogError("dataPath is not exist:" + dataPath);
                return default;
            }
            
            BaseContainer container =  anialysis.Analysis(dataPath);
            
            return new KeyValuePair<string, BaseContainer>(tableName, container);
            
        }

    }
}