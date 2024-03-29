using Spire.Doc;
using Microsoft.VisualBasic;
using System.Data;
using word = Microsoft.Office.Interop.Word;
using Spire.Pdf.Conversion;
using System.Threading;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Reflection.Emit;
using Spire.Pdf.Tables;

namespace MultipleMailMerger
{
    public partial class Form1 : Form
    {
        private List<string> caminhosDocs = new List<string>();
        private List<string> listaCampos = new List<string>();
        public string tabela = "";
        public string caminhoEscolhido = "";

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
            progBar.Visible = false;
        }

        private void btnEscolherDocs_Click(object sender, EventArgs e)
        {
            //Resetar variaveis para come�ar processo do zero
            caminhosDocs.Clear();
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
                //Adiciona o caminho dos ficheiros para vir a manipular posteriormente a uma lista
                foreach (string caminho in janelaEscolherDocs.FileNames)
                {
                    caminhosDocs.Add(caminho);
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
            dgvDados.Columns[0].ReadOnly = true;
            dgvDados.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
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
            int quantDocs = caminhosDocs.Count();
            int quantCampos;
            string campoAVerificar;
            for (int i = 0; i < quantDocs; i++)
            {
                quantCampos = new Document(caminhosDocs[i]).MailMerge.GetMergeFieldNames().Length;
                if (quantCampos > 0)
                {
                    //Percorrer todos os campos presentes no documento
                    //Adicion�-los a uma lista evitando repeti��es
                    for (int j = 0; j < quantCampos; j++)
                    {
                        campoAVerificar = new Document(caminhosDocs[i]).MailMerge.GetMergeFieldNames()[j];
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
                if (!string.IsNullOrEmpty(string.Format("{0}", dgvDados.SelectedRows[i].Cells[0].Value)))
                {
                    bd.strSQL = $"DELETE FROM {tabela} WHERE rowid = {dgvDados.SelectedRows[i].Cells[0].Value}";
                    bd.ExecutarQuery();
                }
            }
            AtualizarGrid();
        }

        private void btnCriarDocs_Click(object sender, EventArgs e)
        {
            //Prepara a janela para escolher pasta de said
            FolderBrowserDialog pastaSaida = new FolderBrowserDialog();
            pastaSaida.Description = "Escolha uma pasta para guardar os documentos";
            pastaSaida.UseDescriptionForTitle = true;
            pastaSaida.ShowNewFolderButton = true;

            //Executa a Janela
            if (pastaSaida.ShowDialog() == DialogResult.OK)
            {
                //Mostra a barra de progresso
                progBar.Visible = true;

                //Bloqueia a intera��o com o from
                //Enquanto corre o processo em background
                btnEscolherDocs.Enabled = false;
                btnAtualizar.Enabled = false;
                btnGuardar.Enabled = false;
                btnApagar.Enabled = false;
                dgvDados.Enabled = false;
                btnCriarDocs.Enabled = false;

                //atribui o caminho escolhido para saida
                caminhoEscolhido = pastaSaida.SelectedPath;

                //Executa processo em background
                exportadorDocumentos.RunWorkerAsync();
            }
        }

        private void exportadorDocumentos_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            word.Application wordapp = new word.Application();
            string caminhoTemp, caminhoSaida;

            //Algoritmo para fazer update � barra de progresso:
            //Como nao d� para usar decimais, aumentamos o valor m�ximo para os milhares
            progBar.Step = 1000 / (dgvDados.SelectedRows.Count * caminhosDocs.Count);

            //Ciclo para percorrer as rows e criar os documentos para cada row selecionada
            for (int i = 0; i < dgvDados.SelectedRows.Count; i++)
            {
                List<string> conteudo = new List<string>();
                DataGridViewRow row = new DataGridViewRow();
                row = dgvDados.SelectedRows[i];

                foreach (DataGridViewCell cell in row.Cells)
                {
                    conteudo.Add(string.Format("{0}", cell.Value));
                }
                //remove o id que vem da tabela
                //n�o � usado nos campos para os documentos
                conteudo.RemoveAt(0);

                //ciclo para criar os documentos com a informa��o da row especifica
                for (int j = 0; j < caminhosDocs.Count; j++)
                {
                    //Caminho para o local onde ficar� o ficheiro .docx "tempor�rio" populado a manipular
                    caminhoTemp = $"{Path.GetDirectoryName(caminhosDocs[j])}\\{Path.GetFileNameWithoutExtension(caminhosDocs[j])}_temp.docx";
                    //Caminho para o local onde ficar� o ficheiro .pdf tratado e exportado
                    caminhoSaida = $"{caminhoEscolhido}\\{Path.GetFileNameWithoutExtension(caminhosDocs[j])}";

                    //Instancia e carrega o documento atrav�s do caminho absoluto
                    Document document = new Document(caminhosDocs[j]);

                    //Executa o mail merge / popula os campos
                    document.MailMerge.Execute(listaCampos.ToArray(), conteudo.ToArray());
                    document.SaveToFile(caminhoTemp, FileFormat.Docx2019);

                    //Para remover a linha de usar o nuget package Spire.Doc sem licen�a
                    //Editamos o documento e removemos o primeiro paragrafo usando uma biblioteca do sistema
                    //Esta biblioteca faz uso do MS Word do sistema em background para realizar as opera��es (+ lento)
                    word.Document doc = wordapp.Documents.Open(caminhoTemp);
                    doc.Paragraphs[1].Range.Delete();

                    //Verificamos se existem documentos no local a exportar com o mesmo nome
                    //Se existir altera o nome do ficheiro
                    caminhoSaida = AlterarNomeFicheiroCasoExista(caminhoSaida);

                    //Guarda em pdf
                    doc.ExportAsFixedFormat(caminhoSaida, word.WdExportFormat.wdExportFormatPDF);

                    //Fecha o documentos .docx "tempor�rios" e guarda os conteudos
                    doc.Close();
                    //Apaga o ficheiro .docx "tempor�rio" pois j� n�o � mais necess�rio
                    File.Delete(caminhoTemp);

                    //Tenta converter para PDF/A
                    //Vers�o FREE da biblioteca Spire.PDF paga
                    //M�ximo de paginas PDF 10!!!
                    //10 P�GINAS!
                    try
                    {
                        PdfStandardsConverter pdf_a = new PdfStandardsConverter(caminhoSaida);
                        pdf_a.ToPdfA1B(caminhoSaida);

                    }
                    catch (Exception)
                    {
                        //Caso nao der muda o nome para indicar que nao est� como pdf/a
                        //Verifica��o para evitar nomes iguais de ficheiros
                        //Remove a extensao .pdf do caminho absoluto do ficheiro para poder adicionar palavra de aviso que nao � PDF/A
                        File.Move(caminhoSaida, AlterarNomeFicheiroCasoExista(caminhoSaida.Remove(caminhoSaida.LastIndexOf('.')) + " NOT_PDF_A"));
                    }

                    //Incrementa barra de progresso
                    progBar.PerformStep();
                }
            }
            //Fecha MS Word que corre em 2� plano
            wordapp.Quit();
        }

        private void exportadorDocumentos_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Ficheiro(s) exportado(s)");

            // hide progress bar
            progBar.Hide();

            //Reseta a barra de progresso
            progBar.Value = 0;

            //Permite novamente a intera��o com o from
            btnEscolherDocs.Enabled = true;
            btnAtualizar.Enabled = true;
            btnGuardar.Enabled = true;
            btnApagar.Enabled = true;
            dgvDados.Enabled = true;
            btnCriarDocs.Enabled = true;
        }

        /// <summary>
        /// M�todo que recebe o caminho absoluto de um ficheiro SEM A EXTENS�O e verifica se existe ficheiros com nomes iguais nesse local <br></br>
        /// Se existir altera o nome do ficheiro no estilo do copiar e renomear do Windows "filename (1).ext"
        /// </summary>
        /// <param name="caminho"></param>
        /// <returns>Caminho absoluto com o nome do ficheiro alterado</returns>
        private static string AlterarNomeFicheiroCasoExista(string caminho)
        {
            //Algoritmo para mudar o nome do ficheiro a exportar em .pdf
            //Em caso de j� existir evita substitui��o/sobreposi��o
            //Tem que ser feito � m�o pois n�o h� metodos que replicam o processo do Windows usado no sistema operativo.
            int c = 1;
            string copias = "";
            string ext = ".pdf";
            while (File.Exists($"{caminho}{copias}{ext}"))
            {
                copias = $" ({c})";
                c++;
            }
            //Ap�s ser encontrado um nome livre, mont�mos o caminho completo.
            caminho = caminho + copias + ext;

            return caminho;
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
                btnApagar.Show();

                for (int i = 0; i < dgvDados.SelectedRows.Count; i++)
                {
                    //N�o deixa exportar se a linha nao estiver guardada na base de dados
                    //Evita tambem exportar a ultima linha vazia
                    if (string.IsNullOrEmpty(string.Format("{0}", dgvDados.SelectedRows[i].Cells[0].Value)))
                    {
                        btnCriarDocs.Hide();

                        //Caso alguma das rows estiver vazia
                        //sai do ciclo pois nao vale a pena continuar
                        return;
                    }
                    else
                    {
                        btnCriarDocs.Show();
                    }
                }
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

        //N�o permite fechar a aplica��o se o background worker estiver a trabalhar
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (exportadorDocumentos.IsBusy)
            {
                e.Cancel = true;
                MessageBox.Show("N�o � permitido encerrar a aplica��o,\nenquanto o exportador de documentos est� a trabalhar!");
            }
        }
    }
}