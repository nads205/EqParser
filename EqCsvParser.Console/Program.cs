namespace EqCsvParser.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            var reader = new Reader();
            reader.ReadAllAdFiles();
            reader.ReadHrFileNew();
            reader.CompareRecords();
        }
    }
}
