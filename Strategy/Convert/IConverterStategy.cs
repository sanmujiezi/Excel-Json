using System.Collections.Generic;

namespace ExcelConvert.Controller
{
    
    public interface IConverterStategy
    {
        public void Convert(string fileName,BaseContainer container);
    }
}