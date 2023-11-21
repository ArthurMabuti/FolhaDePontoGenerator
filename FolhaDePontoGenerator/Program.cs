using FolhaDePontoGenerator;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System.Drawing;

internal partial class Program
{
    private static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        FileInfo file = new(@"C:\Users\Arthur\Desktop\Excel.xlsx");

        await GerarPlanilha(file);
    }

    private static async Task GerarPlanilha(FileInfo file)
    {
        DeletarArquivoAnterior(file);

        using ExcelPackage package = new(file);

        ExcelWorksheet ws = package.Workbook.Worksheets.Add("FolhaDePonto_Arthur");

        ImprimirCabecalho(ws);
        DateTime dataInicial = DateTime.Parse("26/out");
        ImprimirDatas(ws, dataInicial);
        ImprimirTracos(ws);
        ImprimirFimDeSemana(ws);
        ImprimirHorarios(ws);

        var range = ws.Cells["A1:I31"];
        FormatarCelulas(range);
        await package.SaveAsync();
    }

    private static CelulaExcel GerarCelula(ExcelRange range)
    {
        string[] endereco = range.AddressAbsolute.Split('$');
        char coluna = char.Parse(endereco[1]);
        int linha = int.Parse(endereco[2]);
        CelulaExcel celula = new(coluna, linha);
        return celula;
    }

    private static void ImprimirHorarios(ExcelWorksheet ws)
    {
        ExcelRange rangeDiaSemana = ws.Cells[2, 2];
        while (!string.IsNullOrEmpty(rangeDiaSemana.Value?.ToString()))
        {
            CelulaExcel celulaDiaSemana = GerarCelula(rangeDiaSemana);

            if (!FimDeSemana(rangeDiaSemana))
            {
                ExcelRange diaUtil = ws.Cells[$"{++celulaDiaSemana.Coluna}{celulaDiaSemana.Linha}"];
                CelulaExcel celulaDiaUtil = GerarCelula(diaUtil);

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

    private static void ImprimirDatas(ExcelWorksheet ws, DateTime dataInicial)
    {
        var datas = GerarDatas(dataInicial);

        int row = 2;
        int col = 1;

        foreach (DateTime data in datas)
        {
            ws.Cells[row, col].Value = data.ToString("dd/MMM");
            ws.Cells[row, col + 1].Value = PrimeiraLetraMaiscula(data.ToString("dddd"));
            Console.WriteLine(data.ToString("dd/MMM"));
            row++;
        }
    }

    private static void ImprimirTracos(ExcelWorksheet ws)
    {
        int row = 2;
        int col = 8;

        //Criar método que identifique as últimas colunas e imprima os traços

        while(row != 32)
        {
            ws.Cells[row, col].Value = "-";
            ws.Cells[row, col + 1].Value = "-";
            row++;
        }
    }

    private static void ImprimirCabecalho(ExcelWorksheet ws)
    {
        ws.Cells["A1"].Value = " ";
        ws.Cells["B1"].Value = "Dia";
        ws.Cells["C1"].Value = "Entrada";
        ws.Cells["D1"].Value = "Inicio Almoço";
        ws.Cells["E1"].Value = "Fim Almoço";
        ws.Cells["F1"].Value = "Saída";
        ws.Cells["G1"].Value = "HD";
        ws.Cells["H1"].Value = "Quilometragem";
        ws.Cells["I1"].Value = "Pedágio";

        ws.Cells["A1:I1"].Style.Fill.SetBackground(Color.FromArgb(155, 194, 230));
    }

    private static bool FimDeSemana(ExcelRange range)
    {
        if (range.Value.ToString() == "Sábado" || range.Value.ToString() == "Domingo")
            return true;
        return false;
    }

    private static void ImprimirFimDeSemana(ExcelWorksheet ws)
    {
        int row = 2;
        int col = 2;
        while (!string.IsNullOrEmpty(ws.Cells[row, col].Value?.ToString()))
        {
            if (FimDeSemana(ws.Cells[row, col]))
            {
                while (col != 10)
                {
                    ws.Cells[row, col].Style.Fill.SetBackground(Color.FromArgb(128, 128, 128));
                    col++;
                }
                col = 2;
            }
            row++;
        }
    }

    private static void FormatarCelulas(ExcelRange range)
    {
        range.AutoFitColumns();

        int row = range.Rows;
        int col = range.Columns;

        for(int i = 1; i <= row; i++)
        {
            for(int j = 1; j <= col; j++)
            {
                range[i,j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                range[i,j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }
    }

    private static void DeletarArquivoAnterior(FileInfo file)
    {
        if (file.Exists) file.Delete();
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

    public static string PrimeiraLetraMaiscula(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return string.Empty;
        }
        return $"{input[0].ToString().ToUpper()}{input.Substring(1)}";
    }
}