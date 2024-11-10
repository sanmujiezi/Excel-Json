using UnityEngine;

namespace ExcelConvert.Controller
{
    public class ExcelReadStrategy : IReadStrategy
    {
        public void ReadData(string path, out string modelName, out string containerModelName)
        {
            modelName = CreateMoel();
            containerModelName = CreateContainerModel();
            Debug.Log($"{path}从读取了Excel文件");
        }

        public string CreateMoel()
        {
            return "Model";
        }

        public string CreateContainerModel()
        {
            return "ContainerModel";
        }
    }
}