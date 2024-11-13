using System.Collections.Generic;

namespace ExcelConvert.Controller
{
    public interface IReadStrategy
    {
        public void CreateModel(string path);
        public Dictionary<string,BaseContainer> ReadExcel(string path);
    }
    
}