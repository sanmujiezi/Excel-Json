namespace ExcelConvert.Controller
{
    public class ConvertReader
    {
        IReadStrategy readStrategy;

        public ConvertReader(IReadStrategy readStrategy)
        {
            this.readStrategy = readStrategy;
        }

        public void ReadData(string path, out string modelName, out string containermodelName)
        {
            modelName = "";
            containermodelName = "";
            readStrategy.ReadData(path, out modelName, out containermodelName);
        }
    }
}