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
            btnAtualizar.Text = "Atualizar";
            btnGuardar.Text = "Guardar";
            /*
             * Background Color azul claro
             * Autoscale ou auto size ON
             * Por Icon
             */

            //ESCONDER CONTROLOS
            dgvDados.Hide();
            btnAtualizar.Hide();
            btnGuardar.Hide();

            //PROPRIEDADES DA DATAGRID
        }

        private void btnEscolherDocs_Click(object sender, EventArgs e)
        {
            //Resetar variaveis para começar processo do zero
            caminhos.Clear();
            listaDocumentos.Clear();
            listaCampos.Clear();
            tabela = "";

            OpenFileDialog janelaEscolherDocs = new OpenFileDialog();

            //PROPRIEDADES DA JANELA
            janelaEscolherDocs.Title = "Selecionar documentos a tratar";
            janelaEscolherDocs.Filter = "MS Office Word (*.docx;*.dotx)|*.docx;*.dotx"; //Filtro o tipos de ficheiros a carregar
            janelaEscolherDocs.Multiselect = true;                                      //Ativa a seleção de vários ficheiros
            janelaEscolherDocs.RestoreDirectory = true;                                 //Guarda o ultimo caminho usado (para melhor usabilidade)

            //Executa a caixa de diálogo
            //SÓ HAVENDO SUBMISSÃO É QUE CONTINUA A SEQUENCIA DO PROGRAMA
            if (janelaEscolherDocs.ShowDialog() == DialogResult.OK)
            {
                //Captura o caminho dos ficheiros para vir a manipular posteriormente
                //E carrega os documentos para a lista
                foreach (string caminho in janelaEscolherDocs.FileNames)
                {
                    caminhos.Add(caminho);

                    Document documento = new Document();
                    documento.LoadFromFile(caminho);
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

                        if (!string.IsNullOrEmpty(tabela))
                        {
                            //Nomes tabelas em SQL nao podem conter espaços!
                            tabela = tabela.Replace(' ', '_');
                            //Coloca data atual para evitar nomes de tabelas repetidos
                            tabela = tabela + DateTime.UtcNow.AddHours(1).ToString(@"yyyyMMddhhmmss");

                            dbManager.CriarTabela(tabela, listaCampos);
                        }
                        else
                        {
                            MessageBox.Show("Impossível continuar!\nRefira um nome para a tabela!");
                            return; //Pára a execução do método
                        }
                    }

                    AtualizarGrid();
                    dgvDados.Show();
                    btnAtualizar.Show();
                    btnGuardar.Show();
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
            string campoAVerificar;
            for (int i = 0; i < quantDocs; i++)
            {
                quantCampos = listaDocumentos[i].MailMerge.GetMergeFieldNames().Length;
                if (quantCampos > 0)
                {
                    //Percorrer todos os campos presentes no documento
                    //Adicioná-los a uma lista evitando repetições
                    for (int j = 0; j < quantCampos; j++)
                    {
                        campoAVerificar = listaDocumentos[i].MailMerge.GetMergeFieldNames()[j];
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
            int quantColunas;
            string colunaAVerificar;

            //Carregar nomes das tabelas da bd para memória
            bd.strSQL = "SELECT name FROM sqlite_schema " +
                        "WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' " +
                        "ORDER BY 1;";
            nomesTabelas = bd.ExecutarQuery();

            //Ciclo para percorrer as tabelas
            int quantTabelas = nomesTabelas.Rows.Count;
            for (int i = 0; i < quantTabelas; i++)
            {
                //Carregar nomes das colunas da tabela especifica da bd para memória
                bd.strSQL = "SELECT name " +
                           $"FROM PRAGMA_TABLE_INFO('{(string)nomesTabelas.Rows[i][0]}');";
                nomesColunas = bd.ExecutarQuery();

                //Ciclo para percorrer as colunas
                quantColunas = nomesColunas.Rows.Count;
                for (int j = 0; j < quantColunas; j++)
                {
                    colunaAVerificar = (string)nomesColunas.Rows[j][0];
                    if (listaCampos.Contains(colunaAVerificar))
                    {
                        //ultima interação
                        //significa correspondencia total
                        if (j+1 == quantColunas)
                        {
                            tabela = (string)nomesTabelas.Rows[i][0];
                            return true;
                        }
                    }
                }
            }
            //Caso não haja tabelas correspondentes
            return false;
        }

        private void btnAtualizar_Click(object sender, EventArgs e)
        {
            AtualizarGrid();
        }

        private void AtualizarGrid()
        {
            DbManager bd = new DbManager();
            bd.strSQL = $"SELECT rowid, * FROM {tabela};";
            dgvDados.DataSource = bd.ExecutarQuery();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            string nomesColunas = "rowid, ";
            DbManager bd = new DbManager();

            //Ciclo para montar nomes das colunas
            for (int i = 0; i < listaCampos.Count; i++)
            {
                if (i+1 < listaCampos.Count)
                {
                    nomesColunas += listaCampos[i] + ", ";
                }
                else
                {
                    //ultima iteração, nao leva ", "
                    nomesColunas += listaCampos[i];
                }
            }

            //Ciclo para percorrer linhas
            for (int i = 0; i < dgvDados.Rows.Count - 1; i++)
            {
                //Ciclo para montar dados da linha a inserir (percorrer colunas)
                string conteudo = "";
                for (int j = 0; j < dgvDados.ColumnCount; j++)
                {
                    if (j + 1 < dgvDados.ColumnCount)
                    {
                        conteudo += $"'{dgvDados.Rows[i].Cells[j].Value}', ";
                    }
                    else
                    {
                        //ultima coluna, nao leva ", "
                        conteudo += $"'{dgvDados.Rows[i].Cells[j].Value}'";
                    }
                }

                //Guarda os dados da linha na BaseDados
                //bd.strSQL = $"INSERT INTO {tabela} ({nomesColunas}) VALUES ({conteudo});";
                bd.strSQL = $"INSERT OR REPLACE INTO {tabela} ({nomesColunas}) VALUES ({conteudo});";
                bd.ExecutarQuery();
            }

            //Depois de Guardar, atualiza a grid
            AtualizarGrid();
        }
    }
}