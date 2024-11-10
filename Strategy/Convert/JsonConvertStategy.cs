namespace ExcelConvert.Controller
{
    public class JsonConvertStategy<T> : IConverterStategy<T> where T: class
    {
        public void Convert(string outputPath, T t)
        {
            throw new System.NotImplementedException();
        }
    }
}