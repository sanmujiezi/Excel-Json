using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ExcelConvert.Controller;
using LitJson;
using OfficeOpenXml;
using UnityEditor;
using UnityEngine;


namespace ExcelConvert
{
   
    namespace View
    {
        public class ExcelConvertView : EditorWindow
        {
            private static float windowsWidth = 300f;
            private static float windowsHeight = 200f;
            private int convertType = 0;

            [MenuItem("Tools/Excel Converter")]
            public static void ShowWindow()
            {
                var windows = GetWindow<ExcelConvertView>();
                CenterWindow(windows);
                windows.minSize = new Vector2(windowsWidth + 30f, windowsHeight);
            }

            private static void CenterWindow(ExcelConvertView window)
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

                if (GUILayout.Toggle(convertType == 1,"Json",  GUILayout.Width(100)))
                {
                    convertType = 1;
                    ExcelConvertController.Instance.SetConvertType((ConvertType)convertType);
                }
                if (GUILayout.Toggle(convertType == 2,"Binary",  GUILayout.Width(100)))
                {
                    convertType = 2;
                    ExcelConvertController.Instance.SetConvertType((ConvertType)convertType);
                }
                if (GUILayout.Toggle(convertType == 3,"Xml",  GUILayout.Width(100)))
                {
                    convertType = 3;
                    ExcelConvertController.Instance.SetConvertType((ConvertType)convertType);
                }
                
                GUILayout.EndHorizontal();
                
                GUILayout.BeginHorizontal();
                GUILayout.Label("InputPath: ", GUILayout.Width(80));
                GUILayout.TextField(ExcelConvertController.Instance.excelPath, GUILayout.Width(200));
                if (GUILayout.Button("...", GUILayout.Width(20)))
                {
                    InputPath();
                }

                GUILayout.EndHorizontal();

                GUILayout.BeginHorizontal();
                GUILayout.Label("OutputPath: ", GUILayout.Width(80));
                GUILayout.TextField(ExcelConvertController.Instance.outputPath, GUILayout.Width(200));
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
                ExcelConvertController.Instance.ConvertToJson();
            }


            private void InputPath()
            {
                string temp_excelPath = EditorUtility.OpenFilePanel("Select a File", Application.dataPath + "", "xlsx");
                if (File.Exists(temp_excelPath))
                {
                    Debug.Log(temp_excelPath + "_");
                    ExcelConvertController.Instance.SetExcelPath(temp_excelPath);
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
                    Debug.Log(temp_outputPath + "_");
                    ExcelConvertController.Instance.SetOutputPath(temp_outputPath);
                }
                else
                {
                    Debug.Log("Folder not found");
                }
            }
        }
    }
}