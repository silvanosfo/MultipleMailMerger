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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnEscolherDocs = new System.Windows.Forms.Button();
            this.dgvDados = new System.Windows.Forms.DataGridView();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnAtualizar = new System.Windows.Forms.Button();
            this.btnApagar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDados)).BeginInit();
            this.SuspendLayout();
            // 
            // btnEscolherDocs
            // 
            this.btnEscolherDocs.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnEscolherDocs.BackgroundImage")));
            this.btnEscolherDocs.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnEscolherDocs.Location = new System.Drawing.Point(12, 12);
            this.btnEscolherDocs.Name = "btnEscolherDocs";
            this.btnEscolherDocs.Size = new System.Drawing.Size(44, 44);
            this.btnEscolherDocs.TabIndex = 0;
            this.btnEscolherDocs.UseVisualStyleBackColor = true;
            this.btnEscolherDocs.Click += new System.EventHandler(this.btnEscolherDocs_Click);
            // 
            // dgvDados
            // 
            this.dgvDados.AllowUserToDeleteRows = false;
            this.dgvDados.BackgroundColor = System.Drawing.Color.AliceBlue;
            this.dgvDados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDados.Location = new System.Drawing.Point(10, 62);
            this.dgvDados.Name = "dgvDados";
            this.dgvDados.RowTemplate.Height = 25;
            this.dgvDados.Size = new System.Drawing.Size(759, 324);
            this.dgvDados.TabIndex = 1;
            this.dgvDados.SelectionChanged += new System.EventHandler(this.dgvDados_SelectionChanged);
            // 
            // btnGuardar
            // 
            this.btnGuardar.BackColor = System.Drawing.Color.Transparent;
            this.btnGuardar.BackgroundImage = global::MultipleMailMerger.Properties.Resources.save;
            this.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnGuardar.Location = new System.Drawing.Point(725, 12);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(44, 44);
            this.btnGuardar.TabIndex = 2;
            this.btnGuardar.UseVisualStyleBackColor = false;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // btnAtualizar
            // 
            this.btnAtualizar.BackColor = System.Drawing.Color.Transparent;
            this.btnAtualizar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnAtualizar.BackgroundImage")));
            this.btnAtualizar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAtualizar.Location = new System.Drawing.Point(675, 12);
            this.btnAtualizar.Name = "btnAtualizar";
            this.btnAtualizar.Size = new System.Drawing.Size(44, 44);
            this.btnAtualizar.TabIndex = 3;
            this.btnAtualizar.UseVisualStyleBackColor = false;
            this.btnAtualizar.Click += new System.EventHandler(this.btnAtualizar_Click);
            // 
            // btnApagar
            // 
            this.btnApagar.BackColor = System.Drawing.Color.Transparent;
            this.btnApagar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnApagar.BackgroundImage")));
            this.btnApagar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnApagar.Location = new System.Drawing.Point(725, 392);
            this.btnApagar.Name = "btnApagar";
            this.btnApagar.Size = new System.Drawing.Size(44, 44);
            this.btnApagar.TabIndex = 4;
            this.btnApagar.UseVisualStyleBackColor = false;
            this.btnApagar.Click += new System.EventHandler(this.btnApagar_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(781, 442);
            this.Controls.Add(this.btnApagar);
            this.Controls.Add(this.btnAtualizar);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.dgvDados);
            this.Controls.Add(this.btnEscolherDocs);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dgvDados)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Button btnEscolherDocs;
        private DataGridView dgvDados;
        private Button btnGuardar;
        private Button btnAtualizar;
        private Button btnApagar;
    }
}