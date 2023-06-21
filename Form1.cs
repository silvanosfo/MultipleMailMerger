using System.Windows.Forms;
using Spire.Doc;
using System;
using Microsoft.VisualBasic;
using System.Data;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.Xml.Linq;
using Spire.Doc.Fields.Shapes;
using word = Microsoft.Office.Interop.Word;
using Spire.Pdf.Conversion;

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
            //this.AutoSize = true;
            //this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.StartPosition = FormStartPosition.CenterScreen;
            //this.MaximumSize = new Size(1780, 900);

            //ESCONDER CONTROLOS
            dgvDados.Hide();
            btnAtualizar.Hide();
            btnGuardar.Hide();
            btnApagar.Hide();
            btnCriarDocs.Hide();
        }

        private void btnEscolherDocs_Click(object sender, EventArgs e)
        {
            //Resetar variaveis para come�ar processo do zero
            caminhos.Clear();
            listaDocumentos.Clear();
            listaCampos.Clear();
            tabela = "";

            OpenFileDialog janelaEscolherDocs = new OpenFileDialog();

            //PROPRIEDADES DA JANELA
            janelaEscolherDocs.Title = "Selecionar documentos a tratar";
            janelaEscolherDocs.Filter = "MS Office Word (*.docx;*.dotx)|*.docx;*.dotx"; //Filtro o tipos de ficheiros a carregar
            janelaEscolherDocs.Multiselect = true;                                      //Ativa a sele��o de v�rios ficheiros
            janelaEscolherDocs.RestoreDirectory = true;                                 //Guarda o ultimo caminho usado (para melhor usabilidade)

            //Executa a caixa de di�logo
            //S� HAVENDO SUBMISS�O � QUE CONTINUA A SEQUENCIA DO PROGRAMA
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

                //S� CONTINUA SE OS DOCUMENTOS POSSUIREM CAMPOS
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
                    //Atribui � variavel da classe o nome da tabela a usar que corresponde aos campos dos documentos
                    if (!VerificarCorrespondenciaBD())
                    {
                        tabela = Interaction.InputBox("N�o existem tabelas na base de dados ou n�o est�o " +
                                                      "compat�veis com os campos do(s) documento(s).\n\n" +
                                                      "Vai ser criada uma tabela com os campos do(s)\n" +
                                                      "documento(s) na base de dados.\n" +
                                                      "Qual o nome que quer dar a essa tabela?", "Nome da Tabela");

                        if (!string.IsNullOrEmpty(tabela))
                        {
                            //Nomes tabelas em SQL nao podem conter espa�os!
                            tabela = tabela.Replace(' ', '_');
                            //Coloca data atual para evitar nomes de tabelas repetidos
                            tabela = tabela + DateTime.UtcNow.AddHours(1).ToString(@"yyyyMMddhhmmss");

                            dbManager.CriarTabela(tabela, listaCampos);
                        }
                        else
                        {
                            MessageBox.Show("Imposs�vel continuar!\nRefira um nome para a tabela!");
                            return; //P�ra a execu��o do m�todo
                        }
                    }
                    AtualizarGrid();
                    dgvDados.Show();
                    btnAtualizar.Show();
                    btnGuardar.Show();

                    FormatarDGV();
                }
                else
                {
                    MessageBox.Show("Imposs�vel continuar\n" +
                                    "N�o existem campos no(s) documento(s)!");
                }
            }
        }

        private void FormatarDGV()
        {
            //PROPRIEDADES DA DATAGRID
            dgvDados.Columns [0].ReadOnly = true;
            dgvDados.Columns [0].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            dgvDados.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            //N�o pode ordenar por colunas
            foreach (DataGridViewColumn dgvc in dgvDados.Columns)
            {
                dgvc.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        /// <summary>
        /// Verifica se existem campos nos documentos carregados. <br/>
        /// Se existir carrega os campos para a mem�ria
        /// </summary>
        /// <returns>
        /// <see langword="true"/> se existirem campos, caso contr�rio, <see langword="false"/>
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
                    //Adicion�-los a uma lista evitando repeti��es
                    for (int j = 0; j < quantCampos; j++)
                    {
                        campoAVerificar = listaDocumentos[i].MailMerge.GetMergeFieldNames()[j];
                        if (!listaCampos.Contains(campoAVerificar))
                        {
                            listaCampos.Add(campoAVerificar);
                        }

                        //Ultima itera��o, todos os campos encontrados e adicionados
                        if (j+1 == quantCampos)
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    //Nenhum campo encontrado at� ao ultimo documento
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
        /// Se existir faz a atribui��o do nome da tabela
        /// </summary>
        /// <returns>
        /// <see langword="true"/> se existir correspondencia, caso contr�rio, <see langword="false"/>
        /// </returns>
        private bool VerificarCorrespondenciaBD()
        {
            DbManager bd = new DbManager();
            DataTable nomesTabelas = new DataTable();
            DataTable nomesColunas = new DataTable();
            int quantColunas;
            string colunaAVerificar;

            //Carregar nomes das tabelas da bd para mem�ria
            bd.strSQL = "SELECT name FROM sqlite_schema " +
                        "WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' " +
                        "ORDER BY 1;";
            nomesTabelas = bd.ExecutarQuery();

            //Ciclo para percorrer as tabelas
            int quantTabelas = nomesTabelas.Rows.Count;
            for (int i = 0; i < quantTabelas; i++)
            {
                //Carregar nomes das colunas da tabela especifica da bd para mem�ria
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
                        //ultima intera��o
                        //significa correspondencia total
                        if (j+1 == quantColunas)
                        {
                            tabela = (string)nomesTabelas.Rows[i][0];
                            return true;
                        }
                    }
                }
            }
            //Caso n�o haja tabelas correspondentes
            return false;
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            string nomesColunas = "rowid, ";
            bool retiradoRowid = false;
            DbManager bd = new DbManager();

            ////////////////
            //EDITAR DADOS
            ////////////////
            
            //Ciclo para montar nomes das colunas
            for (int i = 0; i < listaCampos.Count; i++)
            {
                if (i+1 < listaCampos.Count)
                {
                    nomesColunas += listaCampos[i] + ", ";
                }
                else
                {
                    //ultima itera��o, nao leva ", "
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

                ////////////////
                //INSERIR DADOS
                ////////////////

                //Se a primeira coluna (rowid) estiver vazia
                if (string.IsNullOrEmpty(dgvDados.Rows[i].Cells[0].Value.ToString()))
                {
                    //Ao guardar v�rias registos novos, � 2�x j� ter� sido retirado "rowid, " da string
                    //Logo, isto evita retirar outros nomes de colunas da string
                    //S� retira o "rowid, " apenas da primeira vez
                    if (!retiradoRowid)
                    {
                        int pos1 = nomesColunas.IndexOf(' ');
                        nomesColunas = nomesColunas.Remove(0, pos1 + 1);
                        retiradoRowid = true;
                    }
                    int pos2 = conteudo.IndexOf(' ');
                    conteudo = conteudo.Remove(0, pos2 + 1);
                }

                //Guarda os dados da linha na BaseDados
                bd.strSQL = $"INSERT OR REPLACE INTO {tabela} ({nomesColunas}) VALUES ({conteudo});";
                bd.ExecutarQuery();
            }

            //Depois de Guardar, atualiza a grid
            AtualizarGrid();
        }

        private void btnApagar_Click(object sender, EventArgs e)
        {
            DbManager bd = new DbManager();
            for (int i = 0; i < dgvDados.SelectedRows.Count; i++)
            {
                //Evitar erro de apagar ultima linha vazia
                if (dgvDados.SelectedRows[i].Cells[0].Value != null)
                {
                    bd.strSQL = $"DELETE FROM {tabela} WHERE rowid = {dgvDados.SelectedRows[i].Cells[0].Value}";
                    bd.ExecutarQuery();
                }
            }
            AtualizarGrid();
        }

        private void btnCriarDocs_Click(object sender, EventArgs e)
        {
            List<string> conteudo = new List<string>();
            DataGridViewRow row = new DataGridViewRow();
            row = dgvDados.CurrentRow;

            foreach (DataGridViewCell cell in row.Cells)
            {
                conteudo.Add(string.Format("{0}", cell.Value));
            }
            //remove o id que vem da tabela
            //n�o � usado nos campos para os documentos
            conteudo.RemoveAt(0);

            //Prepara a janela para escolher pasta de said
            FolderBrowserDialog pastaSaida = new FolderBrowserDialog();
            pastaSaida.Description = "Escolha uma pasta para guardar os documentos";
            pastaSaida.UseDescriptionForTitle = true;
            pastaSaida.ShowNewFolderButton = true;

            //Executa a Janela
            if (pastaSaida.ShowDialog() == DialogResult.OK)
            {
                word.Application wordapp = new word.Application();
                string caminhoTemp, ext;

                for (int i = 0; i < listaDocumentos.Count; i++)
                {
                    //Caminho para o local onde ficar� o ficheiro .docx "tempor�rio" populado a manipular
                    caminhoTemp = $"{Path.GetDirectoryName(caminhos[i])}\\{Path.GetFileNameWithoutExtension(caminhos[i])}_temp.docx";

                    listaDocumentos[i].MailMerge.Execute(listaCampos.ToArray(), conteudo.ToArray());
                    listaDocumentos[i].SaveToFile(caminhoTemp, FileFormat.Docx2019);

                    //Para remover a linha de usar o nuget package Spire.Doc sem licen�a
                    //Editamos o documento e removemos o primeiro paragrafo usando uma biblioteca do sistema
                    //Esta biblioteca faz uso do MS Word do sistema em background para realizar as opera��es (+ lento)
                    word.Document doc = wordapp.Documents.Open(caminhoTemp);
                    doc.Paragraphs[1].Range.Delete();

                    //Algoritmo para mudar o nome do ficheiro a exportar em .pdf
                    //Em caso de j� existir evita substitui��o/sobreposi��o
                    //Tem que ser feito � m�o pois n�o h� metodos que replicam o processo do Windows usado no sistema operativo.
                    int c = 1;
                    string copias = "";
                    ext = ".pdf";
                    string caminhoSaida = $"{pastaSaida.SelectedPath}\\{Path.GetFileNameWithoutExtension(caminhos[i])}";
                    while (File.Exists($"{caminhoSaida}{copias}{ext}"))
                    {
                        copias = $" ({c})";
                        c++;
                    }

                    //Ap�s ser encontrado um nome livre, monta�mos o caminho completo.
                    caminhoSaida = caminhoSaida + copias + ext;

                    //Guarda em pdf
                    doc.ExportAsFixedFormat(caminhoSaida, word.WdExportFormat.wdExportFormatPDF);

                    //Fecha o documentos .docx "tempor�rios" e guarda os conteudos
                    doc.Close();
                    //Apaga o ficheiro .docx "tempor�rio" pois j� n�o � mais necess�rio
                    File.Delete(caminhoTemp);

                    //Converte o pdf em PDF/A 
                    //Vers�o FREE da biblioteca Spire.PDF paga
                    //M�ximo de paginas PDF 10!!!
                    //10 P�GINAS!
                    PdfStandardsConverter pdf_a = new PdfStandardsConverter(caminhoSaida);
                    pdf_a.ToPdfA1B(caminhoSaida);

                }
                //Fecha MS Word que corre em 2� plano
                wordapp.Quit();

                MessageBox.Show("Ficheiro/s criado/s");
            }
        }

        private void AtualizarGrid()
        {
            DbManager bd = new DbManager();
            bd.strSQL = $"SELECT rowid, * FROM {tabela};";
            dgvDados.DataSource = bd.ExecutarQuery();
        }

        private void btnAtualizar_Click(object sender, EventArgs e)
        {
            AtualizarGrid();
        }

        private void dgvDados_SelectionChanged(object sender, EventArgs e)
        {
            //Quando houver rows selecionadas, ativa o botao de apagar
            if (dgvDados.SelectedRows.Count > 0)
            {
                //Por agora s� d� para criar documentos de um registo de cada vez
                //MUDAR ALGORITMO DE EXPORTAR DOCUMENTOS
                //para permitar exportar documentos para os v�rios registos
                if (dgvDados.SelectedRows.Count == 1)
                {
                    btnCriarDocs.Show();
                }
                else
                {
                    btnCriarDocs.Hide();
                }

                btnApagar.Show();
            }
            else
            {
                btnApagar.Hide();
                btnCriarDocs.Hide();
            }
        }

        private void btnEscolherDocs_MouseHover(object sender, EventArgs e)
        {
            tTipDetails.Show("Carregar documentos", btnEscolherDocs);
        }

        private void btnAtualizar_MouseHover(object sender, EventArgs e)
        {
            tTipDetails.Show("Atualizar", btnAtualizar);
        }

        private void btnGuardar_MouseHover(object sender, EventArgs e)
        {
            tTipDetails.Show("Guardar", btnGuardar);
        }

        private void btnApagar_MouseHover(object sender, EventArgs e)
        {
            tTipDetails.Show("Apagar", btnApagar);
        }

        private void btnCriarDocs_MouseHover(object sender, EventArgs e)
        {
            tTipDetails.Show("Exportar", btnCriarDocs);
        }
    }
}