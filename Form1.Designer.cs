﻿namespace MultipleMailMerger
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
            this.btnEscolherDocs = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnEscolherDocs
            // 
            this.btnEscolherDocs.Location = new System.Drawing.Point(12, 12);
            this.btnEscolherDocs.Name = "btnEscolherDocs";
            this.btnEscolherDocs.Size = new System.Drawing.Size(149, 23);
            this.btnEscolherDocs.TabIndex = 0;
            this.btnEscolherDocs.Text = "button1";
            this.btnEscolherDocs.UseVisualStyleBackColor = true;
            this.btnEscolherDocs.Click += new System.EventHandler(this.btnEscolherDocs_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(872, 468);
            this.Controls.Add(this.btnEscolherDocs);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private Button btnEscolherDocs;
    }
}