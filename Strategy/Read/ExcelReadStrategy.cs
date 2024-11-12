using System.Data;
using System.IO;
using Excel;
using UnityEngine;

namespace ExcelConvert.Controller
{
    public class ExcelReadStrategy : IReadStrategy
    {
        public void ReadData(string path, out string modelName, out string containerModelName)
        {
            modelName = "";
            containerModelName = "";
            using (FileStream fs = new FileStream(path,FileMode.Open,FileAccess.Read))
            {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(fs);
                DataSet result = excelReader.AsDataSet();
                if (result.Tables.Count < 0)
                {
                    Debug.Log($"Excel文件为空");
                    return;
                }
                
                //读取文件中所有的表
                for (int i = 0; i < result.Tables.Count; i++)
                {
                    DataTable table = result.Tables[i];
                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        DataRow row = table.Rows[i];
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            Debug.Log(row[k]);
                        }
                    }
                    
                }
                
                Debug.Log($"从{path}读取了Excel文件");    
            }
            
        }

  
       
    }
}