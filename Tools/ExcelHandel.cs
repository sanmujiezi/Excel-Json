using System;
using System.Data;
using System.IO;
using System.Reflection;
using Excel;
using ExcelConvert;
using Plugins.ExcelConvertToJson.Define;
using UnityEngine;

namespace Plugins.ExcelConvertToJson.Tools
{
    public static class ExcelHandel
    {
        public static readonly int tableTitle = 0;
        public static readonly int tableType = 1;
        public static readonly int tableKey = 2;
        public static readonly int tableValue = 4;

        //判断文件中的表的数据结构类是否存在
        public static bool IsExistClass(string className)
        {
            Type _class = Type.GetType(className);
            return _class != null;
        }

        //创建结构类
        public static void CreateModel(DataTable table, string className)
        {
            string path = PathDefine.ModelPath + className + PostfixDefine.Class;
            string code = @"[System.Serializable]" +
                          "\tpublic class " + className + " : BaseModel {";
            DataRow titleRow = GetTitleRow(table);
            DataRow typeRow = GetTypeRow(table);
            for (int i = 0; i < table.Columns.Count; i++)
            {
                // Debug.Log("类型"+typeRow[i] + "，字段名" + titleRow[i]);
                code += $"\n\tpublic {ConvertType(typeRow[i].ToString())} {titleRow[i]};";
            }

            code += "\n}";
            File.WriteAllText(path, code);

            Debug.Log("写入了文件 " + className + PostfixDefine.Class);
        }

        public static void CreateContainer(string className)
        {
            using (StreamWriter writer = new StreamWriter(PathDefine.ModelPath + className + PostfixDefine.Container + PostfixDefine.Class))
            {
                writer.WriteLine("[System.Serializable]");
                writer.WriteLine("public class " + className + PostfixDefine.Container + " : BaseContainer {");
                writer.WriteLine("}");
            }

            Debug.Log("写入了文件 " + className + PostfixDefine.Container + PostfixDefine.Class);
        }

        public static BaseContainer ReadExcel(DataTable table)
        {
            int keycolumus = GetKeyColumus(table);
            string className = table.TableName;
            string containerName = className + PostfixDefine.Container;
            Type classType = Type.GetType(className);
            Type containerType = Type.GetType(containerName);


            ConstructorInfo classCon = classType.GetConstructor(new Type[0]);
            FieldInfo[] classFields = classType.GetFields();

            ConstructorInfo conCon = containerType.GetConstructor(new Type[0]);
            FieldInfo conFields = containerType.GetField("table");
            MethodInfo addMethod = conFields.FieldType.GetMethod("Add");;

            var container = conCon.Invoke(new object[0]);

            for (int i = tableValue; i < table.Rows.Count; i++)
            {
                DataRow row = table.Rows[i];
                var item = classCon.Invoke(new object[0]);
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    Type typeInfo = classFields[j].FieldType;
                    object converValue = Convert.ChangeType(row[j], typeInfo);
                    classFields[j].SetValue(item, converValue);
                }

                addMethod.Invoke(conFields, new object[] { row[keycolumus].ToString(), (item as BaseModel) });
            }

            return container as BaseContainer;
        }


        public static string[] GetExcelFileNames(string path)
        {
            string[] names;
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                DataSet dataSet = reader.AsDataSet();
                names = new string[dataSet.Tables.Count];
                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    names[i] = dataSet.Tables[i].TableName;
                }
            }

            return names;
        }

        //获取标题行
        public static DataRow GetTitleRow(DataTable table)
        {
            return table.Rows[tableTitle];
        }

        //获取类型行
        public static DataRow GetTypeRow(DataTable table)
        {
            return table.Rows[tableType];
        }

        //获取键行
        public static DataRow GetKeyRow(DataTable table)
        {
            return table.Rows[tableKey];
        }

        public static int GetKeyColumus(DataTable table)
        {
            DataRow keyRow = GetKeyRow(table);
            for (int i = 0; i < table.Columns.Count; i++)
            {
                if (keyRow[i].ToString() == "key")
                {
                    return i;
                }
            }

            Debug.LogError($"为{table.TableName}表中不存在 主键 key");
            return -1;
        }

        public static string ConvertType(string type)
        {
            switch (type)
            {
                case "string":
                    return "string";
                case "float":
                    return "float";
                case "int":
                    return "int";
                case "bool":
                    return "bool";
                default:
                    return "string";
            }
        }
    }
}