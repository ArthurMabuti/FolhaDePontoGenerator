using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;

namespace FolhaDePontoGenerator.Modelos;
internal class Planilha
{
    public static async Task GerarPlanilha(FileInfo file)
    {
        DeletarArquivoAnterior(file);

        using ExcelPackage package = new(file);

        ExcelWorksheet ws = package.Workbook.Worksheets.Add("FolhaDePonto_Arthur");

        ImprimirCabecalho(ws);
        Data.ImprimirDatas(ws, Data.EscolherDataInicial());
        Data.ImprimirHorarios(ws);
        DataExtraordinaria.ImprimirFimDeSemana(ws);
        DataExtraordinaria.ImprimirFolga(ws);
        DataExtraordinaria.ImprimirHoraExtra(ws);
        ImprimirTracos(ws);

        var range = ws.Cells[$"A1:I{CountLinhas(ws)}"];
        FormatarCelulas(range);

        await package.SaveAsync();
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
    private static void ImprimirTracos(ExcelWorksheet ws)
    {
        int row = 2;
        int col = 8;

        //Criar método que identifique as últimas colunas e imprima os traços

        while (row <= CountLinhas(ws))
        {
            ws.Cells[row, col].Value = "-";
            ws.Cells[row, col + 1].Value = "-";
            row++;
        }
    }
    private static void FormatarCelulas(ExcelRange range)
    {
        range.AutoFitColumns();

        int row = range.Rows;
        int col = range.Columns;

        for (int i = 1; i <= row; i++)
        {
            for (int j = 1; j <= col; j++)
            {
                range[i, j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                range[i, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }
    }
    private static void DeletarArquivoAnterior(FileInfo file)
    {
        if (file.Exists) file.Delete();
    }
    private static int CountLinhas(ExcelWorksheet ws)
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
