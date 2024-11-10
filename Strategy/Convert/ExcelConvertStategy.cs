namespace ExcelConvert.Controller
{
    public class ExcelConvertStategy<T> : IConverterStategy<T> where T : class
    {
        public void Convert(string outputPath, T t)
        {
            
        }
    }
}