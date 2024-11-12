namespace ExcelConvert.Controller
{
    public interface IReadStrategy
    {
        public void ReadData(string path,out string modelName, out string containerModelName);
     
    }
    
}