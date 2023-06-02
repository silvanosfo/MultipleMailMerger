using System.Windows.Forms;
using Spire.Doc;
using System;


namespace MultipleMailMerger
{
    public partial class Form1 : Form
    {
        public List<string> docPaths = new List<string>();
        public List<string> campos = new List<string>();


        public string pathsTest = "";

        public Form1()
        {
            InitializeComponent();
            this.Text = "Simple Data Inserter Into Word Documents";
            btnChooseFiles.Text = "Selecionar documentos";
        }

        private void btnChooseFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog chooseDocs = new OpenFileDialog();
            chooseDocs.Title = "Selecionar documentos a tratar";

            //Filtro de tipos de ficheiros a manipular
            chooseDocs.Filter = "Microsoft Word Templates (*.dotx)|*.dotx|Microsoft Word Documents (*.docx)|*.docx";

            //Ativa a sele��o de v�rios ficheiros
            chooseDocs.Multiselect = true;
            
            //Guarda o ultimo caminho usado (para melhor usabilidade)
            chooseDocs.RestoreDirectory = true;

            //Executa a caixa de di�logo
            //Se houver submiss�o executa a��o
            if (chooseDocs.ShowDialog() == DialogResult.OK)
            {
                //Captura o caminho dos ficheiros para vir a manipular posteriormente
                foreach (string paths in chooseDocs.FileNames)
                {
                    docPaths.Add(paths);
                    pathsTest += paths + "\n";
                }
                MessageBox.Show(pathsTest);
                pathsTest = "";



                //Instancia��o de v�riavel para manipular um documento Word
                Document document = new Document();

                //Carrega o documento
                document.LoadFromFile(chooseDocs.FileNames[0]);

                //Percorrer todos os campos presentes no documento
                //Adicion�-los a uma lista evitando repeti��es
                foreach (var item in document.MailMerge.GetMergeFieldNames())
                {
                    if (!campos.Contains(item))
                    {
                        campos.Add(item);
                    }
                }
            }


            foreach (string campo in campos)
            {
                pathsTest += campo + "\n";
            }
            MessageBox.Show(pathsTest);
            


        }
    }
}