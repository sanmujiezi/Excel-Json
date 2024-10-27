using System.Collections;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using UnityEngine;
using Application = UnityEngine.Device.Application;

public class EpplusTest : MonoBehaviour
{
    private static string path = Application.dataPath + "/Resources/Data/Excel/Data.xlsx";

    void Start()
    {
        //Debug.Log(path);
        if (!File.Exists(path))
        {
            return;
        }

        FileInfo fileInfo = new FileInfo(path);
        using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];

            Debug.Log($"行：{worksheet.Dimension.Rows} 列：{worksheet.Dimension.Columns}");
            for (int i = 2; i <= worksheet.Dimension.Rows; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.Columns; j++)
                {
                    Debug.Log(worksheet.Cells[i,j].Value);
                    Debug.LogWarning(worksheet.Cells[i,j].Value.GetType());
                }
            }
        }
    }

    // Update is called once per frame
    void Update()
    {
    }
}