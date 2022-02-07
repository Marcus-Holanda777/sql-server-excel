using System;
using System.Data;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Conexao conn = new Conexao("cosmos", "cosmos_v14b");

            string query = "select top 100 * from kardex_filial";
            string caminho = @"c:\arquivo\Kardex.xlsx";
            string nomeTabela = "Kardex";
            string nomePlan = "Resumo";

            DataTable tbl = conn.Tabela(nomeTabela, query);
            ExcelExportar excel = new ExcelExportar(caminho, tbl);
            excel.CriaExcel(nomePlan);
        }

    }
}
