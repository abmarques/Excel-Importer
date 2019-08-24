using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Data.SqlClient;
using System.Configuration;

namespace ImportadorExcel {
    class Program {
        #region Attributes
        /*
        public int Id { get; set; }
        public String Status { get; set; }
        public String Auditoria { get; set; }
        public String Codigo_Externo { get; set; }
        public int Empresa_Id { get; set; }
        public String Descricao { get; set; }
        public double Valor { get; set; }
        public int Tipo { get; set; }
        public int Use_Variavel { get; set; }
        */
        #endregion

        //conexão com o Banco
        #region ConexaoBanco
        SqlConnection sqlCon = null;
        private String strCon = ConfigurationManager.ConnectionStrings["gestor2"].ConnectionString;
        private String strSql = String.Empty;

        
        public void openSqlCon() {
            sqlCon = new SqlConnection(strCon);
            sqlCon.Open();
        }
        #endregion

        static void Main(string[] args) {
            Program p = new Program();
            p.openSqlCon();
            OleDbConnection conexao = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\suporte\Downloads\taxa.xlsx;Extended Properties='Excel 12.0 Xml; HDR=YES';");
        

            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from[Sheet1$]", conexao);
            DataSet ds = new DataSet();

            try {
                conexao.Open();
                

                adapter.Fill(ds);
                foreach (DataRow linha in ds.Tables[0].Rows) {
                    Console.WriteLine("id: {0} – status: {1} – auditoria: {2} - empresa_id: {3} - descricao: {4} - valor: {5} - tipo: {5}",
                        linha["id"].ToString(), linha["status"].ToString(), linha["auditoria"].ToString(), linha["empresa_id"].ToString(),
                        linha["descricao"].ToString(), linha["valor"].ToString(), linha["tipo"].ToString());
                }

            } catch (Exception ex) {
                Console.WriteLine("Erro ao acessar os dados: " +ex.Message);
            }
            DataTable dtImport = ds.Tables[0];
            

            Console.ReadLine();
            
        }
    }
}

