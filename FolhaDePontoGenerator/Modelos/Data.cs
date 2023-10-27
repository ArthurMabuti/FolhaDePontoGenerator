namespace FolhaDePontoGenerator.Modelos;
internal class Data
{
    public DateTime DiaMes { get; set; }
    public string? DiaDaSemana => DiaMes.ToString("dddd");

    public List<DateTime> GerarDatas(DateTime dataInicial)
    {
        DateTime novaData = CorrigirAno(dataInicial);
        List<DateTime> datas = new();
        DateTime dataFinal = novaData.AddMonths(1);
        while (novaData != dataFinal)
        {
            datas.Add(novaData);
            novaData = novaData.AddDays(1);
        }
        return datas;
    }

    private DateTime CorrigirAno(DateTime dataInicial)
    {
        DateTime novaData = DateTime.Parse($"{dataInicial.Day}/{dataInicial.Month}/{DateTime.Now.Year}");
        return novaData;
    }
}
