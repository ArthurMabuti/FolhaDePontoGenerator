using FolhaDePontoGenerator.Modelos;
using OfficeOpenXml;

internal partial class Program
{
    private static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        FileInfo file = new(@"C:\Users\Arthur\Desktop\Folha de Ponto - Arthur Mabuti Pereira.xlsx");
        //FileInfo file = new(@"C:\Users\pereira.arthur.ext\Desktop\Excel.xlsx");

        await Planilha.GerarPlanilha(file);
    }
}