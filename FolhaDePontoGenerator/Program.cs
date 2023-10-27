using OfficeOpenXml;

internal partial class Program
{
    private static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        FileInfo file = new(@"C:\Users\Desktop\Excel.xlsx");

        //DateTime dataInicial = DateTime.Parse("25/out");
        //var datas = GerarDatas(dataInicial);

        //foreach(DateTime data in datas)
        //{
        //    Console.WriteLine(data.ToString("dd/MMM/yyyy"));
        //}
    }

    private static List<FolhaDePonto> GetSetupData()
    {
        List<FolhaDePonto> output = new()
        {
        new() 
        };
        return output;
    }
}