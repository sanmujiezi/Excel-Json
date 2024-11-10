namespace ExcelConvert.Controller
{
    public class Converter<T> where T : class
    {
        IConverterStategy<T> strategy;

        public Converter(IConverterStategy<T> strategy)
        {
            this.strategy = strategy;
        }

        public void Convert(string path, T data)
        {
            strategy.Convert(path, data);
        }
    }
}