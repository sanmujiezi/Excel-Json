using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using LitJson;
using OfficeOpenXml;
using UnityEditor;
using UnityEngine;


namespace Json.ExcelToJson
{
    public class ExcelTransformer : EditorWindow
    {
        private static float windowsWidth = 300f;
        private static float windowsHeight = 200f;
        private string excelPath = "";
        private string outputPath = "";

        [MenuItem("Tools/Excel Converter")]
        public static void ShowWindow()
        {
            var windows = GetWindow<ExcelTransformer>();
            CenterWindow(windows);
            windows.minSize = new Vector2(windowsWidth + 30f, windowsHeight);
        }

        private static void CenterWindow(ExcelTransformer window)
        {
            Rect mainWindowRect = GetMainEditorWindowPosition();
            if (mainWindowRect == Rect.zero)
            {
                Debug.LogWarning("Failed to get the main editor window position. Default positioning applied.");
                return;
            }

            float centeredX = mainWindowRect.x + (mainWindowRect.width - windowsWidth) / 2f;
            float centeredY = mainWindowRect.y + (mainWindowRect.height - windowsHeight) / 2f;

            window.position = new Rect(centeredX, centeredY, windowsWidth, windowsHeight);
        }

        private static Rect GetMainEditorWindowPosition()
        {
            Type containerWindowType = typeof(Editor).Assembly.GetType("UnityEditor.ContainerWindow");
            if (containerWindowType == null)
            {
                Debug.LogError("Unable to find UnityEditor.ContainerWindow type.");
                return Rect.zero;
            }

            FieldInfo showModeField =
                containerWindowType.GetField("m_ShowMode", BindingFlags.NonPublic | BindingFlags.Instance);
            PropertyInfo positionProperty =
                containerWindowType.GetProperty("position", BindingFlags.Public | BindingFlags.Instance);

            if (showModeField == null || positionProperty == null)
            {
                Debug.LogError("Unable to retrieve required fields or properties.");
                return Rect.zero;
            }

            foreach (UnityEngine.Object window in Resources.FindObjectsOfTypeAll(containerWindowType))
            {
                int showMode = (int)showModeField.GetValue(window);
                // ShowMode 4 is typically the main Unity editor window
                if (showMode == 4)
                {
                    return (Rect)positionProperty.GetValue(window, null);
                }
            }

            Debug.LogError("Unable to find main editor window.");
            return Rect.zero;
        }

        public void OnGUI()
        {
            GUILayout.BeginArea(new Rect(0, 0, position.width, position.height));
            GUILayout.BeginVertical();
            GUILayout.Label("ExcelConvertToJson", EditorStyles.boldLabel);
            GUILayout.Label("");

            GUILayout.BeginHorizontal();

            GUILayout.Label("InputPath: ", GUILayout.Width(80));
            GUILayout.TextField(excelPath, GUILayout.Width(200));
            if (GUILayout.Button("...", GUILayout.Width(20)))
            {
                InputPath();
            }

            GUILayout.EndHorizontal();

            GUILayout.BeginHorizontal();
            GUILayout.Label("OutputPath: ", GUILayout.Width(80));
            GUILayout.TextField(outputPath, GUILayout.Width(200));
            if (GUILayout.Button("...", GUILayout.Width(20)))
            {
                OutputPath();
            }

            GUILayout.EndHorizontal();

            GUILayout.BeginHorizontal();
            if (GUILayout.Button("Convert", GUILayout.Width(100)))
            {
                ConvertButton();
                //Debug.Log("Convert");
            }


            GUILayout.EndHorizontal();

            GUILayout.EndVertical();
            GUILayout.EndArea();
        }

        private void ConvertButton()
        {
            if (excelPath == null || outputPath == null || excelPath == "" || outputPath == "")
            {
                Debug.LogWarning("》》请选择 输入文件 或者 输出地址 !");
                return;
            }

            ConvertToJson(excelPath, outputPath);
        }


        private void InputPath()
        {
            string temp_excelPath = EditorUtility.OpenFilePanel("Select a File", Application.dataPath + "", "xlsx");
            if (File.Exists(temp_excelPath))
            {
                excelPath = temp_excelPath;
                Debug.Log(excelPath + "_");
            }
            else
            {
                Debug.Log("File not found");
            }
        }

        private void OutputPath()
        {
            string temp_outputPath = EditorUtility.OpenFolderPanel("Select a File", Application.dataPath + "", "");
            if (Directory.Exists(temp_outputPath))
            {
                outputPath = temp_outputPath;
                Debug.Log(outputPath + "_");
            }
            else
            {
                Debug.Log("Folder not found");
            }
        }

        public void ConvertToJson(string inputPath, string outputPath)
        {
            if (!File.Exists(inputPath))
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
    }
}