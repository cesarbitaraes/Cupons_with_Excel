namespace Cupons;

public class Cupons
{
    static int _numeroCupomGlobal = 1;

    public static List<Tuple<string, int>> GerarNumerosCupons(List<Tuple<string, int>> cuponsLidos)
    {
        var cuponsGerados = new List<Tuple<string, int>>();
        foreach (var cupom in cuponsLidos)
        {
            for (int i = 0; i < cupom.Item2; i++)
            {
                cuponsGerados.Add(Tuple.Create(cupom.Item1, _numeroCupomGlobal++));
            }
        }
        return cuponsGerados;
    }
}