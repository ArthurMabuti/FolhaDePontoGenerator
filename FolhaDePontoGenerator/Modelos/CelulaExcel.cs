using OfficeOpenXml;
using System.Drawing;

namespace FolhaDePontoGenerator.Modelos;
internal class CelulaExcel
{
    public char Coluna { get; set; }
    public int Linha { get; set; }

    public CelulaExcel(char coluna, int linha)
    {
        Coluna = coluna;
        Linha = linha;
    }

    public static CelulaExcel GerarCelula(ExcelRange range)
    {
        string[] endereco = range.AddressAbsolute.Split('$');
        char coluna = char.Parse(endereco[1]);
        int linha = int.Parse(endereco[2]);
        CelulaExcel celula = new(coluna, linha);
        return celula;
    }

    public static void PintarCelulaDeCinza(ExcelWorksheet ws, CelulaExcel celula, bool folga)
    {
        if (folga)
        {
            char colunaInicial = celula.Coluna;
            celula.Coluna = 'B';
            while (celula.Coluna <= 'I')
            {
                if (celula.Coluna != 'B')
                    ws.Cells[$"{celula}"].Value = "";
                ws.Cells[$"{celula}"].Style.Fill.SetBackground(Color.FromArgb(128, 128, 128));
                celula.Coluna++;
            }
            celula.Coluna = colunaInicial;
        }
        celula.Linha++;
    }

    public override string? ToString()
    {
        return $"{Coluna}{Linha}";
    }
}
