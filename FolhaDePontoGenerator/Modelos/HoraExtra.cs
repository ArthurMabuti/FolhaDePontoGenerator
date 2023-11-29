using System.Text.RegularExpressions;

namespace FolhaDePontoGenerator.Modelos;
public class HoraExtra
{
    public string? DiaTrabalhado { get; set; }
    public string QtdHoras { get; set; }

    public HoraExtra(string? diaTrabalhado, string qtdHoras)
    {
        DiaTrabalhado = diaTrabalhado;
        QtdHoras = qtdHoras;
    }

    public static string IncrementarHoraExtra(string horaSaida, string horaExtra)
    {
        string[] tempoSaida = horaSaida.Split(':');
        int horaDeSaida = int.Parse(tempoSaida[0]);
        int minutoDeSaida = int.Parse(tempoSaida[1]);

        int[] tempoExtra = SeparadorHoraMinuto(horaExtra);

        int horaIncrementada = horaDeSaida + tempoExtra[0];
        int minutoIncrementado = minutoDeSaida + tempoExtra[1];
        if (minutoIncrementado >= 60)
        {
            horaIncrementada++;
            minutoIncrementado -= 60;
        }
        return $"{horaIncrementada}:{minutoIncrementado.ToString("D2")}";
    }

    private static int[] SeparadorHoraMinuto(string qtdHoras)
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


    private static bool HoraEMinuto(string? texto) => Regex.IsMatch(texto!, @"(\d{2}|\d{1})[hH]\d{2}(min|Min|MIN)");
    private static bool ApenasHora(string? texto) => Regex.IsMatch(texto!, @"(\d{2}|\d{1})[hH]");
    private static bool ApenasMinuto(string? texto) => Regex.IsMatch(texto!, @"\d{2}(min|Min|MIN)");
}
