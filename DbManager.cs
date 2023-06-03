using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultipleMailMerger
{
    public class DbManager
    {
        private static string pasta = "Databases";
        private static string nomeDB = "Database";
        private static string strConn = $"Data Source={pasta}\\{nomeDB}.db;Version=3";
        public string strSQL = "";

        public void ExecutaQuery(string query)
        {
            SQLiteConnection sqliteConnection = new SQLiteConnection(strConn);
            sqliteConnection.Open();
            SQLiteCommand comando = new SQLiteCommand(query, sqliteConnection);
            comando.ExecuteNonQuery();
            sqliteConnection.Close();
        }

        public void CriarDb()
        {
            if (!System.IO.Directory.Exists(pasta))
            {
                System.IO.Directory.CreateDirectory(pasta);
            }

            try
            {
                SQLiteConnection.CreateFile($"{pasta}\\{nomeDB}.db");
            }
            catch (Exception)
            {
                //Não criou a DB por alguma razao (Faltava criar o diretório)
                throw;
            }
        }

        public void CriarTabela(string nomeTabela, List<string> campos)
        {
            //var e ciclo que vai montar os campos da tabela
            string caracteristicas = "";

            for (int i = 0; i < campos.Count(); i++)
            {
                //se não é ultima iteração
                if (i+1 < campos.Count())
                {
                    caracteristicas += campos[i] + " TEXT, ";

                }
                //ultima iteração, nao leva ","
                else
                {
                    caracteristicas += campos[i] + " TEXT";
                }
            }

            strSQL = $"CREATE TABLE IF NOT EXISTS {nomeTabela}({caracteristicas});";
            ExecutaQuery(strSQL);
        }


        /* FALTA FAZER:
         * - Método para ver todas as tabelas existentes e listar os nomes delas para uma listbox
         * 
         */

    }
}
