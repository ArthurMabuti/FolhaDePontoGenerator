using OfficeOpenXml;

namespace FolhaDePontoGenerator.Modelos;
internal class DataExtraordinaria : Data
{
    public static void ImprimirFimDeSemana(ExcelWorksheet ws)
    {
        CelulaExcel celula = new('B', 2);
        while (!string.IsNullOrEmpty(ws.Cells[$"{celula}"].Value?.ToString()))
        {
            CelulaExcel.PintarCelulaDeCinza(ws, celula, FimDeSemana(ws.Cells[$"{celula}"]));
        }
    }
    public static void ImprimirHoraExtra(ExcelWorksheet ws)
    {
        ExcelRange rangeDia = ws.Cells[2, 1];
        List<HoraExtra> listaDeHorasExtras = ListaDeHoraExtra();
        foreach (var horaExtra in listaDeHorasExtras)
        {
            while (!string.IsNullOrEmpty(rangeDia.Value?.ToString()))
            {
                CelulaExcel celulaDia = CelulaExcel.GerarCelula(rangeDia);

                if (horaExtra.DiaTrabalhado == rangeDia.Value.ToString())
                {
                    string horaSaida = ws.Cells[$"F{celulaDia.Linha}"].Value.ToString()!;
                    string horaTotal = HoraExtra.IncrementarHoraExtra(horaSaida!, horaExtra.QtdHoras);
                    ws.Cells[$"F{celulaDia.Linha}"].Value = horaTotal;
                    break;
                }
                rangeDia = ws.Cells[$"{celulaDia.Coluna}{++celulaDia.Linha}"];
            }
        }
    }
    public static void ImprimirFolga(ExcelWorksheet ws)
    {
        CelulaExcel celulaFolga = new('A', 2);
        string[] diasNaoTrabalhados = DiasNaoTrabalhados()!;
        if (diasNaoTrabalhados.Length != 0)
        {
            while (!string.IsNullOrEmpty(ws.Cells[$"{celulaFolga}"].Value?.ToString()))
            {
                CelulaExcel.PintarCelulaDeCinza(ws, celulaFolga, Feriado(ws.Cells[$"{celulaFolga}"], diasNaoTrabalhados));
            }
        }
    }
    private static string[] ListaDeDias(bool afirmativo)
    {
        if (afirmativo)
        {
            Console.WriteLine("Informe os dias seguindo o exemplo ao lado. (Ex: 15/nov, 20/nov, etc)");
            string[] listaDeDias = Console.ReadLine()!.Split(", ");

            return listaDeDias;
        }
        string[] vazio = Array.Empty<string>();
        return vazio;
    }
    private static string[]? DiasNaoTrabalhados()
    {
        Console.WriteLine("Houve folgas ou feriados no período trabalhado? Excluindo fins de semana. (S - Sim | N - Não)");
        string houveFeriado = Console.ReadLine()!;

        return ListaDeDias(houveFeriado.ToLower() == "s" ? true : false);
    }
    private static string[]? DiasComHorasExtras()
    {
        Console.WriteLine("Foi realizado horas extras no período trabalhado? (S - Sim | N - Não)");
        string houveHoraExtra = Console.ReadLine()!;

        return ListaDeDias(houveHoraExtra.ToLower() == "s" ? true : false);
    }
    private static List<HoraExtra> ListaDeHoraExtra()
    {
        List<HoraExtra> listaDeHorasExtras = new();
        string[] dias = DiasComHorasExtras()!;
        for (int i = 0; i < dias.Length; i++)
        {
            Console.WriteLine($"Quantas horas extras foram realizadas no dia {dias[i]}? (Ex: 1h, 30min, 4h15min)");
            string qtdHoraExtra = Console.ReadLine()!;
            HoraExtra novaHoraExtra = new(dias[i], qtdHoraExtra);
            listaDeHorasExtras.Add(novaHoraExtra);
        }
        return listaDeHorasExtras;
    }
    private static bool Feriado(ExcelRange range, string[] feriados)
    {
        foreach (string feriado in feriados)
            if (range.Value?.ToString() == feriado)
                return true;
        return false;
    }
}
