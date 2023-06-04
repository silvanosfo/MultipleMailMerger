using System;
using System.Collections.Generic;
using System.Data;
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
        private static string caminhoDB = $"{pasta}\\{nomeDB}.db";
        private static string strConn = $"Data Source={pasta}\\{nomeDB}.db;Version=3";
        public string strSQL = "";

        public DataTable ExecutarQuery()
        {
            SQLiteConnection sqliteConnection = new SQLiteConnection(strConn);
            sqliteConnection.Open();
            SQLiteCommand comando = new SQLiteCommand(strSQL, sqliteConnection);
            comando.ExecuteNonQuery();

            SQLiteDataAdapter da = new SQLiteDataAdapter(comando);
            DataTable dt = new DataTable();
            da.Fill(dt);

            sqliteConnection.Close();
            return dt;
        }

        public void CriarDb()
        {
            if (!System.IO.Directory.Exists(pasta))
            {
                System.IO.Directory.CreateDirectory(pasta);
            }

            //Para nao substituir a bd com um ficheiro vazio
            if (!System.IO.File.Exists(caminhoDB))
            {
                try
                {
                    SQLiteConnection.CreateFile(caminhoDB);
                }
                catch (Exception)
                {
                    //Não criou a DB por alguma razao (Faltava criar o diretório)
                    throw;
                }
            }
        }

        public void CriarTabela(string nomeTabela, List<string> campos)
        {
            //var e ciclo que vai montar os campos da tabela
            string caracteristicas = "";
            int quantCampos = campos.Count();
            for (int i = 0; i < quantCampos; i++)
            {
                if (i+1 < quantCampos)
                {
                    caracteristicas += campos[i] + " TEXT, ";

                }
                else
                {
                    //ultima iteração, nao leva ","
                    caracteristicas += campos[i] + " TEXT";
                }
            }

            strSQL = $"CREATE TABLE IF NOT EXISTS {nomeTabela}({caracteristicas});";
            ExecutarQuery();
        }
    }
}
