namespace Sinav_Olusturma
{
    partial class Form5
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form5));
            this.lblABaslik = new System.Windows.Forms.Label();
            this.lblBBaslik = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblACevaplar = new System.Windows.Forms.Label();
            this.lblBCevaplar = new System.Windows.Forms.Label();
            this.lblSoruPuani = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnExcelAktar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // lblABaslik
            // 
            this.lblABaslik.AutoSize = true;
            this.lblABaslik.Location = new System.Drawing.Point(26, 21);
            this.lblABaslik.Name = "lblABaslik";
            this.lblABaslik.Size = new System.Drawing.Size(16, 13);
            this.lblABaslik.TabIndex = 0;
            this.lblABaslik.Text = "...";
            // 
            // lblBBaslik
            // 
            this.lblBBaslik.AutoSize = true;
            this.lblBBaslik.Location = new System.Drawing.Point(26, 50);
            this.lblBBaslik.Name = "lblBBaslik";
            this.lblBBaslik.Size = new System.Drawing.Size(16, 13);
            this.lblBBaslik.TabIndex = 0;
            this.lblBBaslik.Text = "...";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(122, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Soru başına düşen puan";
            // 
            // lblACevaplar
            // 
            this.lblACevaplar.AutoSize = true;
            this.lblACevaplar.Location = new System.Drawing.Point(217, 21);
            this.lblACevaplar.Name = "lblACevaplar";
            this.lblACevaplar.Size = new System.Drawing.Size(16, 13);
            this.lblACevaplar.TabIndex = 0;
            this.lblACevaplar.Text = "...";
            // 
            // lblBCevaplar
            // 
            this.lblBCevaplar.AutoSize = true;
            this.lblBCevaplar.Location = new System.Drawing.Point(217, 50);
            this.lblBCevaplar.Name = "lblBCevaplar";
            this.lblBCevaplar.Size = new System.Drawing.Size(16, 13);
            this.lblBCevaplar.TabIndex = 0;
            this.lblBCevaplar.Text = "...";
            // 
            // lblSoruPuani
            // 
            this.lblSoruPuani.AutoSize = true;
            this.lblSoruPuani.Location = new System.Drawing.Point(217, 80);
            this.lblSoruPuani.Name = "lblSoruPuani";
            this.lblSoruPuani.Size = new System.Drawing.Size(16, 13);
            this.lblSoruPuani.TabIndex = 0;
            this.lblSoruPuani.Text = "...";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6,
            this.Column7,
            this.Column8});
            this.dataGridView1.Location = new System.Drawing.Point(29, 159);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(861, 276);
            this.dataGridView1.TabIndex = 1;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Öğrenci Numarası";
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Öğrenci Adı";
            this.Column2.Name = "Column2";
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Sınav Türü";
            this.Column3.Name = "Column3";
            // 
            // Column4
            // 
            this.Column4.HeaderText = "Grup";
            this.Column4.Name = "Column4";
            // 
            // Column5
            // 
            this.Column5.HeaderText = "Doğru";
            this.Column5.Name = "Column5";
            // 
            // Column6
            // 
            this.Column6.HeaderText = "Yanlış";
            this.Column6.Name = "Column6";
            // 
            // Column7
            // 
            this.Column7.HeaderText = "Boş";
            this.Column7.Name = "Column7";
            // 
            // Column8
            // 
            this.Column8.HeaderText = "Not";
            this.Column8.Name = "Column8";
            // 
            // btnExcelAktar
            // 
            this.btnExcelAktar.Location = new System.Drawing.Point(29, 113);
            this.btnExcelAktar.Name = "btnExcelAktar";
            this.btnExcelAktar.Size = new System.Drawing.Size(165, 23);
            this.btnExcelAktar.TabIndex = 2;
            this.btnExcelAktar.Text = "Excel\'e Aktar";
            this.btnExcelAktar.UseVisualStyleBackColor = true;
            this.btnExcelAktar.Click += new System.EventHandler(this.btnExcelAktar_Click);
            // 
            // Form5
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(927, 447);
            this.Controls.Add(this.btnExcelAktar);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.lblSoruPuani);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblBCevaplar);
            this.Controls.Add(this.lblBBaslik);
            this.Controls.Add(this.lblACevaplar);
            this.Controls.Add(this.lblABaslik);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form5";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sonuç Okuma";
            this.Load += new System.EventHandler(this.Form5_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form5_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form5_DragEnter);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblABaslik;
        private System.Windows.Forms.Label lblBBaslik;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblACevaplar;
        private System.Windows.Forms.Label lblBCevaplar;
        private System.Windows.Forms.Label lblSoruPuani;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column7;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column8;
        private System.Windows.Forms.Button btnExcelAktar;
    }
}