using System.Text.RegularExpressions;

namespace FolhaDePontoGenerator;
internal class HoraExtra
{
    public string? DiaTrabalhado { get; set; }
    public string QtdHoras { get; set; }

    public HoraExtra(string? diaTrabalhado, string qtdHoras)
    {
        DiaTrabalhado = diaTrabalhado;
        QtdHoras = qtdHoras;
    }

    public static string IncrementarHora(string horaSaida, string horaExtra)
    {
        return "";
    }

    private int[] SeparadorHoraMinuto(string qtdHoras)
    {
        if (HoraEMinuto(qtdHoras))
        {
            string[] numeros = Regex.Split(qtdHoras, @"\D+");
            int hora = int.Parse(numeros[0]);
            int minuto = int.Parse(numeros[1]);
            return new int[] { hora, minuto };
        }
        if (ApenasHora(qtdHoras))
        {
            string[] numeros = Regex.Split(qtdHoras, @"\D+");
            int hora = int.Parse(numeros[0]);
            int minuto = 0;
            return new int[] { hora, minuto };
        }
        if (ApenasMinuto(qtdHoras))
        {
            string[] numeros = Regex.Split(qtdHoras, @"\D+");
            int hora = 0;
            int minuto = int.Parse(numeros[0]);
            return new int[] { hora, minuto };
        }
        return new int[] { 0, 0 };
    }


    private bool HoraEMinuto(string? texto) => Regex.IsMatch(texto!, @"(\d{2}|\d{1})[hH]\d{2}(min|Min|MIN)");
    private bool ApenasHora(string? texto) => Regex.IsMatch(texto!, @"(\d{2}|\d{1})[hH]");
    private bool ApenasMinuto(string? texto) => Regex.IsMatch(texto!, @"\d{2}(min|Min|MIN)");
}
