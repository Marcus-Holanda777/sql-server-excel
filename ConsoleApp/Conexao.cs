using System;
using System.Data.SqlClient;
using System.Data;

namespace ConsoleApp
{
    class Conexao
    {
        public string Servidor { get; set; }
        public string Banco { get; set; }

        public Conexao()
        {

        }
        public Conexao(string servidor, string banco)
        {
            Servidor = servidor;
            Banco = banco;
        }
        public string GetDns()
        {
            SqlConnectionStringBuilder contexto = new SqlConnectionStringBuilder
            {
                ApplicationName = "Sql Server",
                DataSource = Servidor,
                IntegratedSecurity = true,
                InitialCatalog = Banco,
            };

            return contexto.ToString();
        }
        public DataTable Tabela(string nome, string consulta)
        {
            DataTable tbl = new DataTable(nome);

            using (SqlConnection conn = new SqlConnection(GetDns()))
            {
                using(SqlCommand cmd = new SqlCommand(consulta, conn))
                {
                    conn.Open();

                    SqlDataReader rds = cmd.ExecuteReader();
                    tbl.Load(rds);

                    return tbl;
                }
            }
        }
    }
}