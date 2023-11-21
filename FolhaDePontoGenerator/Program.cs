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
        //FileInfo file = new(@"C:\Users\Arthur\Desktop\Excel.xlsx");
        FileInfo file = new(@"C:\Users\pereira.arthur.ext\Desktop\Excel.xlsx");

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
        ImprimirFolga(ws);
        ImprimirHorarios(ws);

        var range = ws.Cells[$"A1:I{CountLinhas(ws)}"];
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

            Console.WriteLine(rangeDiaSemana.StyleName.ToString());

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
            row++;
        }
    }

    private static void ImprimirTracos(ExcelWorksheet ws)
    {
        int row = 2;
        int col = 8;

        //Criar método que identifique as últimas colunas e imprima os traços

        while(row <= CountLinhas(ws))
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
        CelulaExcel celula = new('B', 2);
        while (!string.IsNullOrEmpty(ws.Cells[$"{celula}"].Value?.ToString()))
        {
            PintarCelulaDeCinza(ws, celula, FimDeSemana(ws.Cells[$"{celula}"]));
        }
    }

    private static void PintarCelulaDeCinza(ExcelWorksheet ws, CelulaExcel celula, bool folga)
    {
        if (folga)
        {
            char colunaInicial = celula.Coluna;
            celula.Coluna = 'B';
            while (celula.Coluna <= 'I')
            {
                ws.Cells[$"{celula}"].Style.Fill.SetBackground(Color.FromArgb(128, 128, 128));
                celula.Coluna++;
            }
            celula.Coluna = colunaInicial;
        }
        celula.Linha++;
    }
    
    private static bool Feriado(ExcelRange range, string[] feriados)
    {
        foreach(string feriado in feriados)
            if (range.Value?.ToString() == feriado)
                return true;
        return false;
    }

    private static string[]? DiasNaoTrabalhados()
    {
        Console.WriteLine("Houve folgas ou feriados no período trabalhado? Excluindo fins de semana. (S - Sim | N - Não)");
        string houveFeriado = Console.ReadLine()!;

        if (houveFeriado.ToLower() == "s")
        {
            Console.WriteLine("Quais dias não foram trabalhados? (Ex: 15/nov, 20/nov, etc)");
            string[] diasNaoTrabalhados = Console.ReadLine()!.Split(", ");

            return diasNaoTrabalhados;
        }
        string[] vazio = Array.Empty<string>();
        return vazio;
    }

    private static void ImprimirFolga(ExcelWorksheet ws)
    {
        CelulaExcel celulaFolga = new('A', 2);
        if(DiasNaoTrabalhados()!.Length != 0)
        {
            while (!string.IsNullOrEmpty(ws.Cells[$"{celulaFolga}"].Value?.ToString()))
            {
                PintarCelulaDeCinza(ws, celulaFolga, Feriado(ws.Cells[$"{celulaFolga}"], DiasNaoTrabalhados()!));
            }
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

    public static int CountLinhas(ExcelWorksheet ws)
    {
        int row = 2;
        int col = 1;
        int contadorLinhas = 1;

        while (!string.IsNullOrEmpty(ws.Cells[row, col].Value?.ToString()))
        {
            row++;
            contadorLinhas++;
        }
        return contadorLinhas;
    }

}