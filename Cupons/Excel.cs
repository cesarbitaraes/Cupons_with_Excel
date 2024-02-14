using OfficeOpenXml;

namespace Cupons;

public class Excel
{
    private const string FilePath = "Cupons.xlsx";

    public static List<Tuple<string, int>>? LerDoArquivo()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        List<Tuple<string, int>>? listaNomesValores = new List<Tuple<string, int>>();

        try
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            
                Console.WriteLine("Dados lidos do arquivo Cupons.xlsx:");
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var nome = worksheet.Cells[row, 3].Text;
                    var valor = int.Parse(worksheet.Cells[row, 4].Text);
                    Console.WriteLine(nome + " - " + valor);
                    Tuple.Create(nome, valor);
                    listaNomesValores.Add(Tuple.Create(nome, valor));
                }
                return listaNomesValores;
            }
        }
        catch (IndexOutOfRangeException e)
        {
            Console.WriteLine("Arquivo Cupons.xlsx não encontrado. Verifique se esse arquivo encontra-se na mesma pasta do arquivo executável.");
            return null;
        }
        catch (Exception e)
        {
            return listaNomesValores;
        }
    }

    public static void EscreverNoArquivo(List<Tuple<string, int>> cupons)
    {
        try
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                var row = 2;
                foreach (var cupom in cupons)
                {
                    worksheet.Cells[row, 5].Value = cupom.Item1;
                    worksheet.Cells[row, 6].Value = cupom.Item2;
                    row++;
                }
                package.Save();
            }
            Console.WriteLine("\nArquivo alterado com sucesso.");
            Console.ReadLine();
        }
        catch (Exception e)
        {
            Console.WriteLine("\nOcorreu o seguinte erro: \n");
            Console.WriteLine(e);
            Console.WriteLine("\nVerifique se o arquivo está aberto enquanto executa o programa. Em caso afirmativo feche-o e execute novamente.");
            Console.ReadLine();
        }
    }
}