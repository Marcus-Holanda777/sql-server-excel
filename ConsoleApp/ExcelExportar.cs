using System;
using ClosedXML.Excel;
using System.Data;

namespace ConsoleApp
{
    class ExcelExportar
    {
        public string Caminho { get; set; }
        public DataTable Tabela { get; set; }

        public ExcelExportar(string caminho, DataTable tabela)
        {
            Caminho = caminho;
            Tabela = tabela;
        }
        public void CriaExcel(string nomePlan)
        {
            using (XLWorkbook pasta = new XLWorkbook())
            {
                IXLWorksheet planilha = pasta.Worksheets.Add(nomePlan);

                int col = 1;
                foreach (DataColumn c in Tabela.Columns)
                {
                    planilha.Cell(1, col).Value = c.ColumnName;
                    col++;
                }

                int linhas = 2;
                foreach (DataRow linha in Tabela.Rows)
                {
                    int colunas = 1;
                    foreach (DataColumn c in Tabela.Columns)
                    {
                        planilha.Cell(linhas, colunas).Value = linha[c.ColumnName];
                        colunas++;
                    }
                    linhas++;
                }

                /* FORMATAR HEADER */
                int ultima_coluna = Tabela.Columns.Count;
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
                pasta.SaveAs(Caminho);
            }
        }
    }
}