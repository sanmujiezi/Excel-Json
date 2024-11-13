using System.Collections.Generic;

namespace ExcelConvert.Controller
{
    
    public interface IConverterStategy
    {
        public void Convert(string outputPath,Dictionary<string,BaseModel> container);
    }
}