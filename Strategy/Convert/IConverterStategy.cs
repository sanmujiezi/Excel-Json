namespace ExcelConvert.Controller
{
    
    public interface IConverterStategy<T>
    {
        public void Convert(string outputPath,T t);
    }
}