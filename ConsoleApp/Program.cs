using System;
using System.Data;
using System.IO;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Conexao conn = new Conexao("servidor", "banco");

            string query = "seu_select";
            string caminho = @"saida_arquivo";
            string nomeTabela = "Nome_da_tabela";
            string nomePlan = "Nome_da_plan";

            DataTable tbl = conn.Tabela(nomeTabela, query);
            ExcelExportar excel = new ExcelExportar(caminho, tbl);
            excel.CriaExcel(nomePlan);
        }

    }
}
