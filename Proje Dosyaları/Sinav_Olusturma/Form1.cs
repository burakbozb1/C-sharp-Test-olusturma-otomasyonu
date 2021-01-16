using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace Sinav_Olusturma
{
    public partial class Form1 : Form
    {
        //SqlConnection cnn = new SqlConnection("Data Source=PAVILLION\\SQLEXPRESS;Initial Catalog=Sinav_Olusturma;Integrated Security = SSPI;");
        //Çalışan
        //SqlConnection cnn = new SqlConnection("Data Source=(LocalDB)\\v11.0;AttachDbFilename=|DataDirectory|\\Sinav_Olusturma.mdf;Integrated Security=True");
        //SqlConnection cnn = new SqlConnection(@"Data Source=DESKTOP-DENLCA9\SQLEXPRESS;Initial Catalog=Sinav_Olusturma;Integrated Security=SSPI;");
        SqlConnection cnn = new SqlConnection(@"Data Source = Burak\SQLEXPRESS;Initial Catalog = Sinav_Olusturma; Integrated Security = True");
        
        public static string user = "";
        public void baglanti_Ac()
        {
            try
            {
                cnn.Open();
            }
            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
                cnn.Close();
                cnn.Open();
            }
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pnlGiris.Top = ((this.Height - pnlGiris.Height) / 2);
            pnlGiris.Left = ((this.Width - pnlGiris.Width) / 2);
            //pcbLogo.ImageLocation = "..\\Resimler\\marmara-logo.jpg";
            lblGiris.Text = "Marmara Üniversitesi\nTest Oluşturma Programı";
            lblGiris.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            label1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            label2.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnGiris.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            baglanti_Ac();
            if (Directory.Exists("dersler"))
            {

            }
            else
            {
                Directory.CreateDirectory("dersler");
            }
        }

        private void btnGiris_Click(object sender, EventArgs e)
        {
            string kullanici, sifre;
            kullanici = txtKadi.Text;
            sifre = txtSifre.Text;
            SqlCommand cmdkullaniciGetir = new SqlCommand();
            cmdkullaniciGetir.Connection = cnn;
            cmdkullaniciGetir.CommandText = "SELECT * FROM Kullanicilar WHERE kullanici='" + kullanici + "' AND sifre ='" + sifre + "'";
            DataTable dataKullanicilar = new DataTable();
            SqlDataAdapter adpKullanicilar = new SqlDataAdapter(cmdkullaniciGetir);
            adpKullanicilar.Fill(dataKullanicilar);
            if (dataKullanicilar.Rows.Count > 0)
            {
                user = kullanici;
                Form2 frm2 = new Form2();
                frm2.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Yanlış Giriş");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                cnn.Close();
            }
            catch
            {

            }
            Application.Exit();
        }

        private void pnlGiris_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
