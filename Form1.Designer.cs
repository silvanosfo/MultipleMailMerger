namespace MultipleMailMerger
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnEscolherDocs = new System.Windows.Forms.Button();
            this.dgvDados = new System.Windows.Forms.DataGridView();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnAtualizar = new System.Windows.Forms.Button();
            this.btnApagar = new System.Windows.Forms.Button();
            this.btnCriarDocs = new System.Windows.Forms.Button();
            this.tTipDetails = new System.Windows.Forms.ToolTip(this.components);
            this.progBar = new System.Windows.Forms.ProgressBar();
            this.exportadorDocumentos = new System.ComponentModel.BackgroundWorker();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDados)).BeginInit();
            this.SuspendLayout();
            // 
            // btnEscolherDocs
            // 
            this.btnEscolherDocs.BackgroundImage = global::MultipleMailMerger.Properties.Resources.load;
            this.btnEscolherDocs.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnEscolherDocs.Location = new System.Drawing.Point(10, 12);
            this.btnEscolherDocs.Name = "btnEscolherDocs";
            this.btnEscolherDocs.Size = new System.Drawing.Size(44, 44);
            this.btnEscolherDocs.TabIndex = 0;
            this.btnEscolherDocs.UseVisualStyleBackColor = true;
            this.btnEscolherDocs.Click += new System.EventHandler(this.btnEscolherDocs_Click);
            this.btnEscolherDocs.MouseHover += new System.EventHandler(this.btnEscolherDocs_MouseHover);
            // 
            // dgvDados
            // 
            this.dgvDados.AllowUserToDeleteRows = false;
            this.dgvDados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvDados.BackgroundColor = System.Drawing.Color.AliceBlue;
            this.dgvDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDados.Location = new System.Drawing.Point(10, 62);
            this.dgvDados.Name = "dgvDados";
            this.dgvDados.RowTemplate.Height = 25;
            this.dgvDados.Size = new System.Drawing.Size(862, 343);
            this.dgvDados.TabIndex = 1;
            this.dgvDados.SelectionChanged += new System.EventHandler(this.dgvDados_SelectionChanged);
            // 
            // btnGuardar
            // 
            this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGuardar.BackColor = System.Drawing.Color.Transparent;
            this.btnGuardar.BackgroundImage = global::MultipleMailMerger.Properties.Resources.save;
            this.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnGuardar.Location = new System.Drawing.Point(828, 12);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(44, 44);
            this.btnGuardar.TabIndex = 2;
            this.btnGuardar.UseVisualStyleBackColor = false;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            this.btnGuardar.MouseHover += new System.EventHandler(this.btnGuardar_MouseHover);
            // 
            // btnAtualizar
            // 
            this.btnAtualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAtualizar.BackColor = System.Drawing.Color.Transparent;
            this.btnAtualizar.BackgroundImage = global::MultipleMailMerger.Properties.Resources.refresh;
            this.btnAtualizar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAtualizar.Location = new System.Drawing.Point(778, 12);
            this.btnAtualizar.Name = "btnAtualizar";
            this.btnAtualizar.Size = new System.Drawing.Size(44, 44);
            this.btnAtualizar.TabIndex = 3;
            this.btnAtualizar.UseVisualStyleBackColor = false;
            this.btnAtualizar.Click += new System.EventHandler(this.btnAtualizar_Click);
            this.btnAtualizar.MouseHover += new System.EventHandler(this.btnAtualizar_MouseHover);
            // 
            // btnApagar
            // 
            this.btnApagar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnApagar.BackColor = System.Drawing.Color.Transparent;
            this.btnApagar.BackgroundImage = global::MultipleMailMerger.Properties.Resources.delete;
            this.btnApagar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnApagar.Location = new System.Drawing.Point(828, 411);
            this.btnApagar.Name = "btnApagar";
            this.btnApagar.Size = new System.Drawing.Size(44, 44);
            this.btnApagar.TabIndex = 4;
            this.btnApagar.UseVisualStyleBackColor = false;
            this.btnApagar.Click += new System.EventHandler(this.btnApagar_Click);
            this.btnApagar.MouseHover += new System.EventHandler(this.btnApagar_MouseHover);
            // 
            // btnCriarDocs
            // 
            this.btnCriarDocs.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCriarDocs.BackColor = System.Drawing.Color.Transparent;
            this.btnCriarDocs.BackgroundImage = global::MultipleMailMerger.Properties.Resources.export;
            this.btnCriarDocs.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnCriarDocs.Location = new System.Drawing.Point(10, 411);
            this.btnCriarDocs.Name = "btnCriarDocs";
            this.btnCriarDocs.Size = new System.Drawing.Size(44, 44);
            this.btnCriarDocs.TabIndex = 5;
            this.btnCriarDocs.UseVisualStyleBackColor = false;
            this.btnCriarDocs.Click += new System.EventHandler(this.btnCriarDocs_Click);
            this.btnCriarDocs.MouseHover += new System.EventHandler(this.btnCriarDocs_MouseHover);
            // 
            // progBar
            // 
            this.progBar.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.progBar.Location = new System.Drawing.Point(325, 419);
            this.progBar.Maximum = 1000;
            this.progBar.Name = "progBar";
            this.progBar.Size = new System.Drawing.Size(250, 28);
            this.progBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progBar.TabIndex = 6;
            // 
            // exportadorDocumentos
            // 
            this.exportadorDocumentos.DoWork += new System.ComponentModel.DoWorkEventHandler(this.exportadorDocumentos_DoWork);
            this.exportadorDocumentos.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.exportadorDocumentos_RunWorkerCompleted);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(884, 461);
            this.Controls.Add(this.progBar);
            this.Controls.Add(this.btnCriarDocs);
            this.Controls.Add(this.btnApagar);
            this.Controls.Add(this.btnAtualizar);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.dgvDados);
            this.Controls.Add(this.btnEscolherDocs);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDados)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Button btnEscolherDocs;
        private DataGridView dgvDados;
        private Button btnGuardar;
        private Button btnAtualizar;
        private Button btnApagar;
        private Button btnCriarDocs;
        private ToolTip tTipDetails;
        private ProgressBar progBar;
        private System.ComponentModel.BackgroundWorker exportadorDocumentos;
    }
}