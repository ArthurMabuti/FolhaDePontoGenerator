using FolhaDePontoGenerator.Modelos;

internal partial class Program
{
    public class FolhaDePonto
    {
        public Data? Data { get; set; }
        public Horario? Horario { get; set; }
        public string? HelpDesk { get; set; }
        public Viagem Viagem { get; set; }
    }
}