namespace FolhaDePontoGenerator;
internal class CelulaExcel
{
    public char Coluna {  get; set; }
    public int Linha { get; set; }

    public CelulaExcel(char coluna, int linha)
    {
        Coluna = coluna;
        Linha = linha;
    }

    public override string? ToString()
    {
        return $"{Coluna}{Linha}";
    }
}
