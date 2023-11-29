using OfficeOpenXml;

namespace FolhaDePontoGenerator.Modelos;
internal class Data
{
    public static DateTime EscolherDataInicial()
    {
        Console.WriteLine("A partir de qual dia inicia o período de trabalho? (Ex: 01/jan, 26/mar)");
        string dia = Console.ReadLine()!;
        DateTime dataInicial = DateTime.Parse(dia);
        return dataInicial;
    }
    public static void ImprimirHorarios(ExcelWorksheet ws)
    {
        ExcelRange rangeDiaSemana = ws.Cells[2, 2];
        while (!string.IsNullOrEmpty(rangeDiaSemana.Value?.ToString()))
        {
            CelulaExcel celulaDiaSemana = CelulaExcel.GerarCelula(rangeDiaSemana);

            if (!FimDeSemana(rangeDiaSemana))
            {
                ExcelRange diaUtil = ws.Cells[$"{++celulaDiaSemana.Coluna}{celulaDiaSemana.Linha}"];
                CelulaExcel celulaDiaUtil = CelulaExcel.GerarCelula(diaUtil);

                diaUtil.Value = "07:30";
                ws.Cells[$"{++celulaDiaUtil.Coluna}{celulaDiaUtil.Linha}"].Value = "12:00";
                ws.Cells[$"{++celulaDiaUtil.Coluna}{celulaDiaUtil.Linha}"].Value = "13:00";
                ws.Cells[$"{++celulaDiaUtil.Coluna}{celulaDiaUtil.Linha}"].Value = "16:30";

                rangeDiaSemana = ws.Cells[$"{--celulaDiaSemana.Coluna}{++celulaDiaSemana.Linha}"];
            }
            else
                rangeDiaSemana = ws.Cells[$"{celulaDiaSemana.Coluna}{++celulaDiaSemana.Linha}"];
        }
    }
    public static void ImprimirDatas(ExcelWorksheet ws, DateTime dataInicial)
    {
        var datas = GerarDatas(dataInicial);

        int row = 2;
        int col = 1;

        foreach (DateTime data in datas)
        {
            ws.Cells[row, col].Value = data.ToString("dd/MMM");
            ws.Cells[row, col + 1].Value = PrimeiraLetraMaiscula(data.ToString("dddd"));
            row++;
        }
    }
    private static List<DateTime> GerarDatas(DateTime dataInicial)
    {
        List<DateTime> datas = new();
        DateTime dataFinal = dataInicial.AddMonths(1);
        while (dataInicial != dataFinal)
        {
            datas.Add(dataInicial);
            dataInicial = dataInicial.AddDays(1);
        }
        return datas;
    }
    public static bool FimDeSemana(ExcelRange range)
    {
        if (range.Value.ToString() == "Sábado" || range.Value.ToString() == "Domingo")
            return true;
        return false;


    }
    private static string PrimeiraLetraMaiscula(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return string.Empty;
        }
        return $"{input[0].ToString().ToUpper()}{input.Substring(1)}";
    }
}
