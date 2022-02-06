using System;
using ClosedXML.Excel;
using System.Data;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Conexao conn = new Conexao("seu_servidor", "seu_banco");

            string query = "seu_select";

            DataTable tbl = conn.Tabela("exemplo", query);
            CriaExcel(@"caminho_do_excel", tbl);
            
        }

        static void CriaExcel(string nome, DataTable dados)
        {
            using (XLWorkbook pasta = new XLWorkbook())
            {
                IXLWorksheet planilha = pasta.Worksheets.Add("nome_planilha");

                int col = 1;
                foreach(DataColumn c in dados.Columns)
                {
                    planilha.Cell(1, col).Value = c.ColumnName;
                    col++;
                }

                int linhas = 2;
                foreach (DataRow linha in dados.Rows)
                {
                    int colunas = 1;
                    foreach (DataColumn c in dados.Columns)
                    {
                        planilha.Cell(linhas, colunas).Value = linha[c.ColumnName];
                        colunas++;
                    }
                    linhas++;
                }

                /* FORMATAR HEADER */
                int ultima_coluna = dados.Columns.Count;
                IXLCell inicio = planilha.Cell(1, 1);
                IXLCell fim = planilha.Cell(1, ultima_coluna);

                IXLRange cabeca = planilha.Range(inicio, fim);

                cabeca.Style
                    .Font.SetBold()
                    .Font.SetFontColor(XLColor.White)
                    .Fill.SetBackgroundColor(XLColor.BlueViolet)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                planilha.SheetView.ZoomScale = 80;
                planilha.Columns().AdjustToContents(); /* ajusta as colunas de forma automatica */
                pasta.SaveAs(nome);
            }
        }
    }
}
