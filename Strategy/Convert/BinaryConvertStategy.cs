using System;
using UnityEngine;

namespace ExcelConvert.Controller
{
    public class BinaryConvertStategy<T> : IConverterStategy<T> where T : class
    {
        public void Convert(string outputPath, T t)
        {
            Debug.Log($"将{typeof(T).Name}转换为而二进制到{outputPath}文件中");
        }
    }
}