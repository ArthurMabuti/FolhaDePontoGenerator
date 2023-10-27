internal partial class Program
{
    public class FolhaDePonto
    {
        public DateTime Data { get; set; }
        public string? DiaDaSemana => Data.ToString("dddd");
        public DateTime Entrada { get; set; }
        public DateTime InicioAlmoco { get; set; }
        public DateTime FimAlmoco { get; set; }
        public DateTime Saida { get; set; }
        public string? HelpDesk { get; set; }
        public int Quilometragem { get; set; }
        public double Pedagio { get; set; }

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
}