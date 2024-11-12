using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using Plugins.ExcelConvertToJson.Define;
using UnityEngine;
using Object = System.Object;

namespace Plugins.ExcelConvertToJson.Tools
{
    public static class ExcelHandel
    {
        public static readonly int tableTitle = 1;
        public static readonly int tableType = 2;
        public static readonly int tableKey = 3;
        public static readonly int tableValue = 5;

        //判断文件中的表的数据结构类是否存在
        public static bool IsExistClass(string className)
        {
            Type _class = Type.GetType(className);
            return _class != null;
        }

        //创建结构类
        public static void CreateModel(DataTable table, string className)
        {
            using (FileStream fs = new FileStream(PathDefine.ModelPath + className + PostfixDefine.Model,
                       FileMode.Create))
            {
                StreamWriter writer = new StreamWriter(fs);
                writer.WriteLine("[System.Serializable]");
                writer.WriteLine("public class " + className + " : BaseModel {");
                DataRow titleRow = GetTitleRow(table);
                DataRow typeRow = GetTypeRow(table);
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    writer.WriteLine($"public {typeRow[i]} {titleRow[i]};");
                }

                writer.WriteLine("}");
            }
        }

        //创建容器
        public static Dictionary<string,T> CreateContainer<T>(DataTable table) where T : new()
        {
            int keyColumus = GetKeyColumus(table);
            if (keyColumus == -1)
            {
                keyColumus = 1;
            }
            
            Dictionary<string,T> container = new Dictionary<string, T>();
            ConstructorInfo constructorInfo = typeof(T).GetConstructor(new Type[0]);
            FieldInfo[] infos = typeof(T).GetFields();

            for (int i = tableValue; i < table.Rows.Count; i++)
            {
                DataRow curRow = table.Rows[i];
                T item = (T)constructorInfo.Invoke(new object[0]);
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    var value = curRow[j];
                    infos[j].SetValue(item, value);
                }
                container.Add(curRow[keyColumus].ToString(), item);
            }

            return container;

        }
        
        public static Dictionary<string,BaseModel> CreateContainer(DataTable table,Type type)
        {
            int keyColumus = GetKeyColumus(table);
            if (keyColumus == -1)
            {
                keyColumus = 1;
            }
            
            Dictionary<string,BaseModel> container = new Dictionary<string, BaseModel>();
            ConstructorInfo constructorInfo = type.GetConstructor(new Type[0]);
            FieldInfo[] infos = type.GetFields();

            for (int i = tableValue; i < table.Rows.Count; i++)
            {
                DataRow curRow = table.Rows[i];
                var item = constructorInfo.Invoke(new object[0]);
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    var value = curRow[j];
                    infos[j].SetValue(item, value);
                }
                container.Add(curRow[keyColumus].ToString(), (BaseModel)item);
            }

            return container;

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
                if (keyRow[i] == "key")
                {
                    return i;
                }
            }
            
            Debug.LogError($"为{table.TableName}表中不存在 主键 key");
            return -1;
        }
    }
}