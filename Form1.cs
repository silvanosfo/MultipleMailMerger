using System.Windows.Forms;
using Spire.Doc;
using System;
using Microsoft.VisualBasic;
using System.Data;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.Xml.Linq;

namespace MultipleMailMerger
{
    public partial class Form1 : Form
    {
        private List<string> caminhos = new List<string>();
        private List<Document> listaDocumentos = new List<Document>();
        private List<string> listaCampos = new List<string>();
        public string tabela = "";

        public Form1()
        {
            InitializeComponent();
            this.Text = "Multiple Mail Merger";
            btnEscolherDocs.Text = "Selecionar documentos";
            /*
             * Background Color azul claro
             * Autoscale ou auto size ON
             * Por Icon
             */
        }

        private void btnEscolherDocs_Click(object sender, EventArgs e)
        {
            OpenFileDialog janelaEscolherDocs = new OpenFileDialog();
            janelaEscolherDocs.Title = "Selecionar documentos a tratar";

            //Filtro de tipos de ficheiros a manipular
            janelaEscolherDocs.Filter = "MS Office Word (*.docx;*.dotx)|*.docx;*.dotx";

            //Ativa a seleção de vários ficheiros
            janelaEscolherDocs.Multiselect = true;
            
            //Guarda o ultimo caminho usado (para melhor usabilidade)
            janelaEscolherDocs.RestoreDirectory = true;

            //Executa a caixa de diálogo
            //SÓ HAVENDO SUBMISSÃO É QUE CONTINUA A SEQUENCIA DO PROGRAMA
            if (janelaEscolherDocs.ShowDialog() == DialogResult.OK)
            {
                //Captura o caminho dos ficheiros para vir a manipular posteriormente
                foreach (string caminho in janelaEscolherDocs.FileNames)
                {
                    caminhos.Add(caminho);
                }

                //Document documento = new Document();
                ///documento.LoadFromFile(janelaEscolherDocs.FileNames[0]);


                foreach (string caminho in janelaEscolherDocs.FileNames)
                {
                    //Instanciação de váriavel para manipular os documentos Word
                    Document documento = new Document();
                    //Carrega os documentos
                    documento.LoadFromFile(caminho);
                    //Adiciona documento carregado à lista
                    listaDocumentos.Add(documento);
                    //VALE A PENA FAZER .Dispose()? Garbage Collector?
                }
                

                //SÓ CONTINUA SE OS DOCUMENTOS POSSUIREM CAMPOS
                if (VerificarExistCamposDocs())
                {
                    //Criar a base de dados
                    DbManager dbManager = new DbManager();
                    try
                    {
                        dbManager.CriarDb();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro: " + ex.Message);
                    }

                    //Antes de usar a variavel tabela nas querys
                    //Fazer isto
                    //Atribui à variavel da classe o nome da tabela a usar que corresponde aos campos dos documentos
                    if (!VerificarCorrespondenciaBD())
                    {
                        tabela = Interaction.InputBox("Não existem tabelas na base de dados ou não estão " +
                                                      "compatíveis com os campos do(s) documento(s).\n\n" +
                                                      "Vai ser criada uma tabela com os campos do(s)\n" +
                                                      "documento(s) na base de dados.\n" +
                                                      "Qual o nome que quer dar a essa tabela?", "Nome da Tabela");

                        //Nomes tabelas em SQL nao podem conter espaços!
                        tabela = tabela.Replace(' ', '_');

                        if (!string.IsNullOrEmpty(tabela))
                        {
                            dbManager.CriarTabela(tabela, listaCampos);
                        }
                        else
                        {
                            MessageBox.Show("Impossível continuar!\nRefira um nome para a tabela!");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Impossível continuar\n" +
                                    "Não existem campos no(s) documento(s)!");
                }
            }
        }

        /// <summary>
        /// Verifica se existem campos nos documentos carregados. <br/>
        /// Se existir carrega os campos para a memória
        /// </summary>
        /// <returns>
        /// <see langword="true"/> se existirem campos, caso contrário, <see langword="false"/>
        /// </returns>
        private bool VerificarExistCamposDocs()
        {
            int quantDocs = listaDocumentos.Count();
            int quantCampos;
            for (int i = 0; i < quantDocs; i++)
            {
                quantCampos = listaDocumentos[i].MailMerge.GetMergeFieldNames().Length;
                if (quantCampos > 0)
                {
                    //Percorrer todos os campos presentes no documento
                    //Adicioná-los a uma lista evitando repetições
                    for (int j = 0; j < quantCampos; j++)
                    {
                        string campoAVerificar = listaDocumentos[i].MailMerge.GetMergeFieldNames()[j];
                        if (!listaCampos.Contains(campoAVerificar))
                        {
                            listaCampos.Add(campoAVerificar);
                        }

                        //Ultima iteração, todos os campos encontrados e adicionados
                        if (j+1 == quantCampos)
                        {
                            return true;
                        }
                    }



                    //foreach (var item in listaDocumentos[i].MailMerge.GetMergeFieldNames())
                    //{
                    //    if (!campos.Contains(item))
                    //    {
                    //        campos.Add(item);

                    //        //Ultima iteração, todos os campos encontrados e adicionados
                    //        if (i+1 == listaDocumentos[i].MailMerge.GetMergeFieldNames().Count())
                    //        {
                    //            return true;
                    //        }
                    //    }
                    //}
                }
                else
                {
                    //Nenhum campo encontrado até ao ultimo documento
                    if (i + 1 == quantDocs)
                    {
                        return false;
                    }
                }
            }
            //Lista vazia
            return false;
        }

        /// <summary>
        /// Verifica se existe na base de dados uma tabela com as colunas iguais aos campos dos documentos. <br/>
        /// Se existir faz a atribuição do nome da tabela
        /// </summary>
        /// <returns>
        /// <see langword="true"/> se existir correspondencia, caso contrário, <see langword="false"/>
        /// </returns>
        private bool VerificarCorrespondenciaBD()
        {
            DbManager bd = new DbManager();
            DataTable nomesTabelas = new DataTable();
            DataTable nomesColunas = new DataTable();

            //Carregar nomes das tabelas da bd para memória
            bd.strSQL = "SELECT name FROM sqlite_schema " +
                        "WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' " +
                        "ORDER BY 1;";
            nomesTabelas = bd.ExecutarQuery();

            //Ciclo para percorrer as tabelas
            for (int i = 0; i < nomesTabelas.Rows.Count; i++)
            {
                //Carregar nomes das colunas da tabela especifica da bd para memória
                bd.strSQL = "SELECT name " +
                           $"FROM PRAGMA_TABLE_INFO('{(string)nomesTabelas.Rows[i][0]}')";
                nomesColunas = bd.ExecutarQuery();

                //Ciclo para percorrer as colunas
                for (int j = 0; j < nomesColunas.Rows.Count; j++)
                {
                    //MUDAR PARA CONTAINS! POIS SEGUIR POR ORDEM ESTÁ ERRADO!
                    //MUDAR PARA CONTAINS! POIS SEGUIR POR ORDEM ESTÁ ERRADO!
                    //MUDAR PARA CONTAINS! POIS SEGUIR POR ORDEM ESTÁ ERRADO!
                    //MUDAR PARA CONTAINS! POIS SEGUIR POR ORDEM ESTÁ ERRADO!
                    //MUDAR PARA CONTAINS! POIS SEGUIR POR ORDEM ESTÁ ERRADO!
                    //MUDAR PARA CONTAINS! POIS SEGUIR POR ORDEM ESTÁ ERRADO!
                    //SE NÃO CONTEM OS NOMES DAS COLUNAS NA LISTA DE CAMPOS NA ULTIMA ITERAÇÃO???
                    //SE NÃO CONTEM OS NOMES DAS COLUNAS NA LISTA DE CAMPOS NA ULTIMA ITERAÇÃO???
                    //SE NÃO CONTEM OS NOMES DAS COLUNAS NA LISTA DE CAMPOS NA ULTIMA ITERAÇÃO???
                    //SE NÃO CONTEM OS NOMES DAS COLUNAS NA LISTA DE CAMPOS NA ULTIMA ITERAÇÃO???
                    if ((string)nomesColunas.Rows[j][0] == listaCampos[j])
                    {
                        //ultima interação
                        //significa correspondencia total
                        if (j+1 == nomesColunas.Rows.Count)
                        {
                            tabela = (string)nomesTabelas.Rows[i][0];
                            return true;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            //Caso não haja tabelas
            return false;
        }
    }
}