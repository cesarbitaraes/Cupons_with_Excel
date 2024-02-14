namespace Cupons;

class Program
{
    private static void Main()
    {
        var cupons = Excel.LerDoArquivo();
        if (cupons is null) return;
        var cuponsGerados = Cupons.GerarNumerosCupons(cupons);
        Excel.EscreverNoArquivo(cuponsGerados);
    }
}