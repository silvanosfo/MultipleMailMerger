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
        private static SQLiteConnection sqliteConnection;

        //Método que retorna a stringconnection usada
        private SQLiteConnection DbConnection()
        {
            sqliteConnection= new SQLiteConnection("Data Source=..\\Databases\\Database.db;Version=3");
            sqliteConnection.Open();
            return sqliteConnection;
        }

        public void CriarDb()
        {
            if (!System.IO.Directory.Exists("..\\Databases"))
            {
                System.IO.Directory.CreateDirectory("..\\Databases");
            }

            try
            {
                SQLiteConnection.CreateFile("..\\Databases\\Database.db");
            }
            catch (Exception)
            {
                //Não criou a DB por alguma razao (Faltava criar o diretório)
                throw;
            }
        }

        public void CriarTabela(string nomeTabela, List<string> campos)
        {
            //var que vai montar os campos da tabela
            string caracteristicas = "id int";
            foreach (var campo in campos)
            {
                caracteristicas += ", " + campo + " TEXT";
            }

            //var do tipo command para executar comandos
            //O QUE RAIO É USING??????
            using (SQLiteCommand cmd = DbConnection().CreateCommand())
            {
                cmd.CommandText = $"CREATE TABLE IF NOT EXISTS {nomeTabela}({caracteristicas});";
                cmd.ExecuteNonQuery();
            }
        }


        /* FALTA FAZER:
         * - Método para ver todas as tabelas existentes e listar os nomes delas para uma listbox
         * 
         */








    }
}
