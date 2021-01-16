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
using System.Reflection;
using word = Microsoft.Office.Interop.Word;

namespace Sinav_Olusturma
{
    public partial class Form2 : Form
    {
        int toplamSayi = 0;
        int rndSeciliId = 0, rndSeciliIDEkle = 0;
        public static int secilenKolaySayi = 0, secilenOrtaSayi = 0, secilenZorSayi = 0;
        public static int secilenToplam = 0;
        public static string dersinAdi;
        public static int dersIDsi;
        public static int kolaySayi, ortaSayi, zorSayi;
        DataTable dtSecilenSorular = new DataTable();
        DataTable dataGelenSorular = new DataTable();
        DataTable rndSecilenSorular = new DataTable();
        DataTable rndTumSorular = new DataTable();
        DataTable rndKalanSorular = new DataTable();
        int ekranYukseklik, ekranGenislik;
        //Çalışan
        //SqlConnection cnn = new SqlConnection("Data Source=(LocalDB)\\v11.0;AttachDbFilename=|DataDirectory|\\Sinav_Olusturma.mdf;Integrated Security=True");
        //SqlConnection cnn = new SqlConnection("Data Source=PAVILLION\\SQLEXPRESS;Initial Catalog=Sinav_Olusturma;Integrated Security = SSPI;");
        SqlConnection cnn = new SqlConnection(@"Data Source = Burak\SQLEXPRESS;Initial Catalog = Sinav_Olusturma; Integrated Security = True");
        int dersID;
        int seciliOzelID;
        int seciliOzelID2;
        string resimYolu = "";
        string aResim = "", bResim = "", cResim = "", dResim = "", eResim = "";
        public static string yazdirilacakTur = "";
        string aYeniIsim = "", bYeniIsim = "", cYeniIsim = "", dYeniIsim = "", eYeniIsim = "", aYeniYol = "", bYeniYol = "", cYeniYol = "", dYeniYol = "", eYeniYol = "";

        public void baglanti_Ac()
        {
            try
            {
                cnn.Open();
            }
            catch
            {
                cnn.Close();
                cnn.Open();
            }
        }
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            //MessageBox.Show(Screen.PrimaryScreen.Bounds.Width.ToString() + " x " + Screen.PrimaryScreen.Bounds.Height.ToString());
            
            ekranYukseklik = Screen.PrimaryScreen.Bounds.Height-62;
            ekranGenislik=Screen.PrimaryScreen.Bounds.Width;

            pnlDersEkle.Visible = false;
            pnlOzelSinav.Visible = false;
            pnlRastgeleOlustur.Visible = false;
            pnlSoruEkle.Visible = false;

            lblHos.Top = 150;
            lblHos.Left = 10;
            

            lblKulAdi.Left = 10;
            lblKulAdi.Top = 180;
            lblKulAdi.Text = Form1.user;

            lblHos.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            lblKulAdi.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            lblHos.BringToFront();
            lblKulAdi.BringToFront();

            //pcbTbmyo.ImageLocation = "tbmyo-logo.jpg";
            pcbTbmyo.BringToFront();
            pcbTbmyo.Left = 20;
            pcbTbmyo.Top = 15;
            //this.BackColor = Color.FromArgb(255, 232, 232);
            pcbSolMenu.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pcbSolMenu.Height = ekranYukseklik;

            pictureBox1.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox2.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox3.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox4.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            pictureBox5.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox6.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox7.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox8.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            pictureBox9.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox10.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox11.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pictureBox12.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            /*pcbBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pcbBaslik.Width = 1280;
            pcbBaslik.Height = 50;
            pcbBaslik.Left = 162;
            pcbBaslik.Top = 0;*/

            //lblSoru.BringToFront();
            btnDersEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSoruEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSinavOlusturPanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            button1.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSonuc.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            btnDersEklePanel.BringToFront();
            btnSoruEklePanel.BringToFront();
            btnSinavOlusturPanel.BringToFront();
            button1.BringToFront();
            btnSonuc.BringToFront();

            btnDersEklePanel.FlatAppearance.BorderSize = 0;
            btnSoruEklePanel.FlatAppearance.BorderSize = 0;
            btnSinavOlusturPanel.FlatAppearance.BorderSize = 0;
            button1.FlatAppearance.BorderSize = 0;
            btnSonuc.FlatAppearance.BorderSize = 0;


            //pcbSolMenu.BackColor = (Color)ColorConverter.ConvertFromString("#F1EFE2")
            pcbOnizleme.ImageLocation = @"resimsiz.jpg";
            label62.Visible = false;
            label63.Visible = false;
            label64.Visible = false;
            label65.Visible = false;
            label78.Visible = false;
            pcbOzlA.Visible = false;
            pcbOzlB.Visible = false;
            pcbOzlC.Visible = false;
            pcbOzlD.Visible = false;
            label58.Visible = false;
            label59.Visible = false;
            label60.Visible = false;
            label61.Visible = false;
            label79.Visible = false;
            pcbRndA.Visible = false;
            pcbRndB.Visible = false;
            pcbRndC.Visible = false;
            pcbRndD.Visible = false;
            pcbRndE.Visible = false;


            btnA.Left = 137;
            btnB.Left = 305;
            btnC.Left = 462;
            btnD.Left = 623;
            btnE.Left = 784;
            btnA.Top = 277;
            btnB.Top = 277;
            btnC.Top = 277;
            btnD.Top = 277;
            btnE.Top = 277;
            rdbSikMetin.Checked = true;
            if (Directory.Exists(@"C:\DersResimleri\"))
            {
                if (Directory.Exists(@"C:\DersResimleri\dersler"))
                {
                    //MessageBox.Show("Klasör Mevcut");
                }
                else
                {
                    Directory.CreateDirectory(@"C:\DersResimleri\dersler");
                    MessageBox.Show("Resimler için klasör oluşturuldu.");
                }

            }
            else
            {
                Directory.CreateDirectory(@"C:\DersResimleri");
                Directory.CreateDirectory(@"C:\DersResimleri\dersler");
                MessageBox.Show("Resimler için klasör oluşturuldu.");
            }

            dtSecilenSorular.Columns.Add("soru_id");
            dtSecilenSorular.Columns.Add("soru");
            dtSecilenSorular.Columns.Add("a_cevap");
            dtSecilenSorular.Columns.Add("b_cevap");
            dtSecilenSorular.Columns.Add("c_cevap");
            dtSecilenSorular.Columns.Add("d_cevap");
            dtSecilenSorular.Columns.Add("e_cevap");
            dtSecilenSorular.Columns.Add("dogru_cevap");
            dtSecilenSorular.Columns.Add("zorluk_derecesi");
            dtSecilenSorular.Columns.Add("soru_resim");

            rndSecilenSorular.Columns.Add("soru_id");
            rndSecilenSorular.Columns.Add("soru");
            rndSecilenSorular.Columns.Add("a_cevap");
            rndSecilenSorular.Columns.Add("b_cevap");
            rndSecilenSorular.Columns.Add("c_cevap");
            rndSecilenSorular.Columns.Add("d_cevap");
            rndSecilenSorular.Columns.Add("e_cevap");
            rndSecilenSorular.Columns.Add("dogru_cevap");
            rndSecilenSorular.Columns.Add("zorluk_derecesi");
            rndSecilenSorular.Columns.Add("soru_resim");

            rndTumSorular.Columns.Add("soru_id");
            rndTumSorular.Columns.Add("soru");
            rndTumSorular.Columns.Add("a_cevap");
            rndTumSorular.Columns.Add("b_cevap");
            rndTumSorular.Columns.Add("c_cevap");
            rndTumSorular.Columns.Add("d_cevap");
            rndTumSorular.Columns.Add("e_cevap");
            rndTumSorular.Columns.Add("dogru_cevap");
            rndTumSorular.Columns.Add("zorluk_derecesi");
            rndTumSorular.Columns.Add("soru_resim");

            rndKalanSorular.Columns.Add("soru_id");
            rndKalanSorular.Columns.Add("soru");
            rndKalanSorular.Columns.Add("a_cevap");
            rndKalanSorular.Columns.Add("b_cevap");
            rndKalanSorular.Columns.Add("c_cevap");
            rndKalanSorular.Columns.Add("d_cevap");
            rndKalanSorular.Columns.Add("e_cevap");
            rndKalanSorular.Columns.Add("dogru_cevap");
            rndKalanSorular.Columns.Add("zorluk_derecesi");
            rndKalanSorular.Columns.Add("soru_resim");

            this.WindowState = FormWindowState.Maximized;
            baglanti_Ac();
            SqlCommand cmdDersleriGetir = new SqlCommand();
            cmdDersleriGetir.Connection = cnn;
            cmdDersleriGetir.CommandText = "SELECT * FROM Dersler";
            DataTable dataDersler = new DataTable();
            SqlDataAdapter adpDersler = new SqlDataAdapter(cmdDersleriGetir);
            adpDersler.Fill(dataDersler);
            for (int i = 0; i < dataDersler.Rows.Count; i++)
            {
                cmbDersAdi.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                cmbSinavDers.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                cmbOzelDersler.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                cmbSilinecekDers.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                cmbOzelDersUst.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
            }
            pcbSoruResmi.ImageLocation = @"resimsiz.jpg";
        }

        private void btnDersEkle_Click(object sender, EventArgs e)
        {
            if (txtDersAdi.Text != "")
            {
                string dersAdi = txtDersAdi.Text;
                SqlCommand cmdDersEkle = new SqlCommand();
                cmdDersEkle.Connection = cnn;
                cmdDersEkle.CommandText = "INSERT INTO Dersler (ders_adi) VALUES ('" + dersAdi + "')";
                try
                {
                    cmdDersEkle.ExecuteNonQuery();
                    if (Directory.Exists(@"C:\DersResimleri\dersler\" + dersAdi))
                    {
                        MessageBox.Show("Klasör mevcut.");
                    }
                    else
                    {
                        Directory.CreateDirectory(@"C:\DersResimleri\dersler\" + dersAdi);
                        MessageBox.Show("Resimler için klasör oluşturuldu.");
                    }
                    MessageBox.Show(dersAdi + " isimli ders eklendi.");

                    //Dersler ilgili comboboxlara yeniden listeleniyor
                    SqlCommand cmdDersleriGetir = new SqlCommand();
                    cmdDersleriGetir.Connection = cnn;
                    cmdDersleriGetir.CommandText = "SELECT * FROM Dersler";
                    DataTable dataDersler = new DataTable();
                    SqlDataAdapter adpDersler = new SqlDataAdapter(cmdDersleriGetir);
                    adpDersler.Fill(dataDersler);
                    cmbDersAdi.Items.Clear();
                    cmbSinavDers.Items.Clear();
                    cmbOzelDersler.Items.Clear();
                    cmbSilinecekDers.Items.Clear();
                    for (int i = 0; i < dataDersler.Rows.Count; i++)
                    {
                        cmbDersAdi.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                        cmbSinavDers.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                        cmbOzelDersler.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                        cmbSilinecekDers.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                    }
                    cmbDersAdi.SelectedItem = null;
                    cmbSinavDers.SelectedItem = null;
                    cmbOzelDersler.SelectedItem = null;
                    cmbSilinecekDers.SelectedItem = null;

                }
                catch (Exception hata)
                {
                    MessageBox.Show("Ders Eklenemedi. Hata= " + hata.Message);
                }
            }
            else
            {
                MessageBox.Show("Lütfen yıldızlı(*) alanları boş bırakmayınız.");
            }


        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void cmbDersAdi_SelectedIndexChanged(object sender, EventArgs e)
        {
            string dersAdi = cmbDersAdi.SelectedItem.ToString();
            SqlCommand cmdDersID = new SqlCommand();
            cmdDersID.Connection = cnn;
            cmdDersID.CommandText = "SELECT ders_id FROM Dersler WHERE ders_adi = '" + dersAdi + "'";
            DataTable dataDersID = new DataTable();
            SqlDataAdapter adpDersID = new SqlDataAdapter(cmdDersID);
            adpDersID.Fill(dataDersID);
            dersID = Convert.ToInt32(dataDersID.Rows[0]["ders_id"]);
        }

        private void btnSoruEkle_Click(object sender, EventArgs e)
        {

            if ((rdbA.Checked == true || rdbB.Checked == true || rdbC.Checked == true || rdbD.Checked == true || rdbE.Checked == true) && rtxtSoru.Text != "" && /*((txtA.Text != "" && txtB.Text != "" && txtC.Text != "" && txtD.Text != "")|| true) &&*/ cmbDersAdi.SelectedItem != null && cmbZorluk.SelectedItem != null)
            {

                string soru = "", cevapA = "", cevapB = "", cevapC = "", cevapD = "", cevapE = "", dogruCevap = "", zorlukDerecesi = "";
                Random rnd = new Random();
                string yeniisim = (rnd.Next(1, 9999999)).ToString();
                string yeniResimYolu;
                if (resimYolu == "")
                {
                    yeniResimYolu = "";
                }
                else
                {
                    yeniResimYolu = @"C:\DersResimleri\dersler\" + cmbDersAdi.SelectedItem.ToString() + @"\" + yeniisim + ".jpg";
                }
                soru = rtxtSoru.Text;
                if (rdbSikMetin.Checked == true)
                {
                    cevapA = txtA.Text;
                    cevapB = txtB.Text;
                    cevapC = txtC.Text;
                    cevapD = txtD.Text;
                    cevapE = txtE.Text;
                }

                else if (rdbSikResim.Checked == true)
                {
                    aYeniIsim = (rnd.Next(1, 9999999)).ToString();
                    bYeniIsim = (rnd.Next(1, 9999999)).ToString();
                    cYeniIsim = (rnd.Next(1, 9999999)).ToString();
                    dYeniIsim = (rnd.Next(1, 9999999)).ToString();
                    eYeniIsim = (rnd.Next(1, 9999999)).ToString();
                    //File.Copy(resimYolu, yeniResimYolu);
                    aYeniYol = @"C:\DersResimleri\dersler\" + cmbDersAdi.SelectedItem.ToString() + @"\" + aYeniIsim + ".jpg";
                    bYeniYol = @"C:\DersResimleri\dersler\" + cmbDersAdi.SelectedItem.ToString() + @"\" + bYeniIsim + ".jpg";
                    cYeniYol = @"C:\DersResimleri\dersler\" + cmbDersAdi.SelectedItem.ToString() + @"\" + cYeniIsim + ".jpg";
                    dYeniYol = @"C:\DersResimleri\dersler\" + cmbDersAdi.SelectedItem.ToString() + @"\" + dYeniIsim + ".jpg";
                    eYeniYol = @"C:\DersResimleri\dersler\" + cmbDersAdi.SelectedItem.ToString() + @"\" + eYeniIsim + ".jpg";
                    cevapA = aYeniYol;
                    cevapB = bYeniYol;
                    cevapC = cYeniYol;
                    cevapD = dYeniYol;
                    cevapE = eYeniYol;
                }

                dogruCevap = "";
                if (rdbA.Checked == true)
                {
                    dogruCevap = cevapA;
                }
                if (rdbB.Checked == true)
                {
                    dogruCevap = cevapB;
                }
                if (rdbC.Checked == true)
                {
                    dogruCevap = cevapC;
                }
                if (rdbD.Checked == true)
                {
                    dogruCevap = cevapD;
                }
                if (rdbE.Checked == true)
                {
                    dogruCevap = cevapE;
                }
                zorlukDerecesi = cmbZorluk.SelectedItem.ToString();
                SqlCommand cmdSoruEkle = new SqlCommand();
                cmdSoruEkle.Connection = cnn;
                try
                {
                    cmdSoruEkle.CommandText = "INSERT INTO Sorular (ders_id,soru,soru_resim,a_cevap,b_cevap,c_cevap,d_cevap,e_cevap,dogru_cevap,zorluk_derecesi) VALUES (@ders_id,@soru,@soru_resim,@a_cevap,@b_cevap,@c_cevap,@d_cevap,@e_cevap,@dogru_cevap,@zorluk_derecesi)";
                    cmdSoruEkle.Parameters.AddWithValue("@ders_id", dersID);
                    cmdSoruEkle.Parameters.AddWithValue("@soru", soru);
                    cmdSoruEkle.Parameters.AddWithValue("@soru_resim", yeniResimYolu);
                    if (rdbSikResim.Checked == true)
                    {
                        File.Copy(aResim, aYeniYol);
                        File.Copy(bResim, bYeniYol);
                        File.Copy(cResim, cYeniYol);
                        File.Copy(dResim, dYeniYol);
                        File.Copy(eResim, eYeniYol);

                    }
                    cmdSoruEkle.Parameters.AddWithValue("@a_cevap", cevapA);
                    cmdSoruEkle.Parameters.AddWithValue("@b_cevap", cevapB);
                    cmdSoruEkle.Parameters.AddWithValue("@c_cevap", cevapC);
                    cmdSoruEkle.Parameters.AddWithValue("@d_cevap", cevapD);
                    cmdSoruEkle.Parameters.AddWithValue("@e_cevap", cevapE);
                    cmdSoruEkle.Parameters.AddWithValue("@dogru_cevap", dogruCevap);
                    cmdSoruEkle.Parameters.AddWithValue("@zorluk_derecesi", zorlukDerecesi);
                    if (resimYolu == "")
                    {
                        yeniResimYolu = "";
                    }

                    else
                    {
                        File.Copy(resimYolu, yeniResimYolu);
                    }
                    //MessageBox.Show(cevapE.ToString());
                    cmdSoruEkle.ExecuteNonQuery();
                    MessageBox.Show("Soru eklendi.");

                    rdbA.Checked = false;
                    rdbB.Checked = false;
                    rdbC.Checked = false;
                    rdbD.Checked = false;
                    rdbE.Checked = false;
                    cmbZorluk.SelectedItem = null;
                    txtA.Text = null;
                    txtB.Text = null;
                    txtC.Text = null;
                    txtD.Text = null;
                    txtE.Text = null;
                    rtxtSoru.Text = null;
                    resimYolu = "";
                    cmbZorluk.SelectedItem = null;
                    pcbSoruResmi.ImageLocation = @"resimsiz.jpg";
                    rtxtOnizleme.Text = null;
                    pcbOnizleme.Image = null;
                    pcbA.Image = null;
                    pcbB.Image = null;
                    pcbC.Image = null;
                    pcbD.Image = null;
                    pcbE.Image = null;
                    aResim = "";
                    bResim = "";
                    cResim = "";
                    dResim = "";
                    eResim = "";
                }
                catch (Exception hata)
                {
                    MessageBox.Show(hata.Message);
                }
            }
            else
            {
                MessageBox.Show("Lütfen yıldızlı(*) alanları boş bırakmayınız.");
            }

        }

        private void btnResimEkle_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            //string dosyayolu = dosya.FileName;
            //MessageBox.Show(dosyayolu);
            resimYolu = dosya.FileName;
            //pcbSoruResmi.ImageLocation = resimYolu;
            if (resimYolu == "")
            {
                pcbOnizleme.ImageLocation = @"resimsiz.jpg";
            }
            else
            {
                pcbOnizleme.ImageLocation = resimYolu;
            }

        }

        private void btnResmiKaldir_Click(object sender, EventArgs e)
        {
            resimYolu = "";
            //pcbSoruResmi.ImageLocation = @"C:\DersResimleri\dersler\resimsiz.jpg";
            pcbOnizleme.ImageLocation = @"resimsiz.jpg";
            //pcbOnizleme.Image = null;
        }

        private void btnSorulariGetir_Click(object sender, EventArgs e)
        {
            kolaySayi = 0;
            ortaSayi = 0;
            zorSayi = 0;
            rndSecilenSorular.Rows.Clear();
            rndTumSorular.Rows.Clear();
            rndKalanSorular.Rows.Clear();


            lstRndGelen.Items.Clear();
            lstRndKalan.Items.Clear();
            if ((cmbSinavDers.SelectedItem.ToString() != "" || cmbSinavDers.SelectedItem != null) && txtSinavKolay.Text != "" && txtSinavOrta.Text != "" && txtSinavZor.Text != "")
            {
                dersinAdi = cmbSinavDers.SelectedItem.ToString();
                SqlCommand cmdDersId = new SqlCommand();
                cmdDersId.Connection = cnn;
                cmdDersId.CommandText = "SELECT ders_id FROM Dersler WHERE ders_adi='" + dersinAdi + "'";
                DataTable dataDersID = new DataTable();
                SqlDataAdapter adpDersID = new SqlDataAdapter(cmdDersId);
                adpDersID.Fill(dataDersID);
                dersIDsi = Convert.ToInt32(dataDersID.Rows[0]["ders_id"]);

                int totalKolay, totalOrta, totalZor;
                //Toplam kolay soru sayısı
                SqlCommand cmdKolaySorular = new SqlCommand();
                cmdKolaySorular.Connection = cnn;
                cmdKolaySorular.CommandText = "SELECT * FROM Sorular WHERE zorluk_derecesi='Kolay' AND ders_id='" + dersIDsi + "'";
                DataTable dataKSoruSayisi = new DataTable();
                SqlDataAdapter adpKSoruSayisi = new SqlDataAdapter(cmdKolaySorular);
                adpKSoruSayisi.Fill(dataKSoruSayisi);
                totalKolay = Convert.ToInt32(dataKSoruSayisi.Rows.Count);

                //Toplam Orta soru sayısı
                SqlCommand cmdOrtaSorular = new SqlCommand();
                cmdOrtaSorular.Connection = cnn;
                cmdOrtaSorular.CommandText = "SELECT * FROM Sorular WHERE zorluk_derecesi='Orta' AND ders_id='" + dersIDsi + "'";
                DataTable dataOSoruSayisi = new DataTable();
                SqlDataAdapter adpOSoruSayisi = new SqlDataAdapter(cmdOrtaSorular);
                adpOSoruSayisi.Fill(dataOSoruSayisi);
                totalOrta = Convert.ToInt32(dataOSoruSayisi.Rows.Count);

                //Toplam Zor soru sayısı
                SqlCommand cmdZorSorular = new SqlCommand();
                cmdZorSorular.Connection = cnn;
                cmdZorSorular.CommandText = "SELECT * FROM Sorular WHERE zorluk_derecesi='Zor' AND ders_id='" + dersIDsi + "'  ";
                DataTable dataZSoruSayisi = new DataTable();
                SqlDataAdapter adpZSoruSayisi = new SqlDataAdapter(cmdZorSorular);
                adpZSoruSayisi.Fill(dataZSoruSayisi);
                totalZor = Convert.ToInt32(dataZSoruSayisi.Rows.Count);




                kolaySayi = Convert.ToInt32(txtSinavKolay.Text);
                ortaSayi = Convert.ToInt32(txtSinavOrta.Text);
                zorSayi = Convert.ToInt32(txtSinavZor.Text);

                if (kolaySayi <= totalKolay && ortaSayi <= totalOrta && zorSayi <= totalZor)
                {
                    /*Form3 frm3 = new Form3();
                    frm3.Show();*/
                    //Seçilen sayı kadar soruların getirileceği bölüm
                    SqlCommand cmdKolayGetir = new SqlCommand();
                    cmdKolayGetir.Connection = cnn;
                    cmdKolayGetir.CommandText = "SELECT * FROM Sorular WHERE zorluk_derecesi='Kolay' AND ders_id='" + dersIDsi + "' ORDER By NEWID()";
                    DataTable dataKolaylar = new DataTable();
                    SqlDataAdapter adpKolaylar = new SqlDataAdapter(cmdKolayGetir);
                    adpKolaylar.Fill(dataKolaylar);

                    SqlCommand cmdOrtaGetir = new SqlCommand();
                    cmdOrtaGetir.Connection = cnn;
                    cmdOrtaGetir.CommandText = "SELECT * FROM Sorular WHERE zorluk_derecesi='Orta' AND ders_id='" + dersIDsi + "' ORDER By NEWID()";
                    DataTable dataOrtalar = new DataTable();
                    SqlDataAdapter adpOrtalar = new SqlDataAdapter(cmdOrtaGetir);
                    adpOrtalar.Fill(dataOrtalar);

                    SqlCommand cmdZorGetir = new SqlCommand();
                    cmdZorGetir.Connection = cnn;
                    cmdZorGetir.CommandText = "SELECT * FROM Sorular WHERE zorluk_derecesi='Zor' AND ders_id='" + dersIDsi + "' ORDER By NEWID()";
                    DataTable dataZorlar = new DataTable();
                    SqlDataAdapter adpZorlar = new SqlDataAdapter(cmdZorGetir);
                    adpZorlar.Fill(dataZorlar);

                    toplamSayi = kolaySayi + ortaSayi + zorSayi;



                    for (int i = 0; i < Form2.kolaySayi; i++)
                    {
                        rndSecilenSorular.Rows.Add(dataKolaylar.Rows[i]["soru_id"].ToString().Trim(), dataKolaylar.Rows[i]["soru"].ToString().Trim(), dataKolaylar.Rows[i]["a_cevap"].ToString().Trim(), dataKolaylar.Rows[i]["b_cevap"].ToString().Trim(), dataKolaylar.Rows[i]["c_cevap"].ToString().Trim(), dataKolaylar.Rows[i]["d_cevap"].ToString().Trim(), dataKolaylar.Rows[i]["e_cevap"].ToString().Trim(), dataKolaylar.Rows[i]["dogru_cevap"].ToString().Trim(), dataKolaylar.Rows[i]["zorluk_derecesi"].ToString().Trim(), dataKolaylar.Rows[i]["soru_resim"].ToString().Trim());
                    }
                    for (int i = 0; i < Form2.ortaSayi; i++)
                    {
                        rndSecilenSorular.Rows.Add(dataOrtalar.Rows[i]["soru_id"].ToString().Trim(), dataOrtalar.Rows[i]["soru"].ToString().Trim(), dataOrtalar.Rows[i]["a_cevap"].ToString().Trim(), dataOrtalar.Rows[i]["b_cevap"].ToString().Trim(), dataOrtalar.Rows[i]["c_cevap"].ToString().Trim(), dataOrtalar.Rows[i]["d_cevap"].ToString().Trim(), dataOrtalar.Rows[i]["e_cevap"].ToString().Trim(), dataOrtalar.Rows[i]["dogru_cevap"].ToString().Trim(), dataOrtalar.Rows[i]["zorluk_derecesi"].ToString().Trim(), dataOrtalar.Rows[i]["soru_resim"].ToString().Trim());
                    }
                    for (int i = 0; i < Form2.zorSayi; i++)
                    {
                        rndSecilenSorular.Rows.Add(dataZorlar.Rows[i]["soru_id"].ToString().Trim(), dataZorlar.Rows[i]["soru"].ToString().Trim(), dataZorlar.Rows[i]["a_cevap"].ToString().Trim(), dataZorlar.Rows[i]["b_cevap"].ToString().Trim(), dataZorlar.Rows[i]["c_cevap"].ToString().Trim(), dataZorlar.Rows[i]["d_cevap"].ToString().Trim(), dataZorlar.Rows[i]["e_cevap"].ToString().Trim(), dataZorlar.Rows[i]["dogru_cevap"].ToString().Trim(), dataZorlar.Rows[i]["zorluk_derecesi"].ToString().Trim(), dataZorlar.Rows[i]["soru_resim"].ToString().Trim());
                    }

                    for (int i = 0; i < rndSecilenSorular.Rows.Count; i++)
                    {
                        lstRndGelen.Items.Add(rndSecilenSorular.Rows[i]["soru_id"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["soru"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["a_cevap"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["b_cevap"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["c_cevap"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["d_cevap"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["e_cevap"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                        lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["soru_resim"].ToString().Trim());
                    }
                    int[] secilenler = new int[rndSecilenSorular.Rows.Count];
                    for (int i = 0; i < rndSecilenSorular.Rows.Count; i++)
                    {
                        secilenler[i] = Convert.ToInt32(rndSecilenSorular.Rows[i]["soru_id"]);
                    }

                    SqlCommand cmdTumSorular = new SqlCommand();
                    cmdTumSorular.Connection = cnn;
                    cmdTumSorular.CommandText = "SELECT * FROM Sorular WHERE ders_id='" + dersIDsi + "'";
                    SqlDataAdapter adpTumSorular = new SqlDataAdapter(cmdTumSorular);
                    adpTumSorular.Fill(rndTumSorular);

                    for (int i = 0; i < rndSecilenSorular.Rows.Count; i++)
                    {
                        for (int j = 0; j < rndTumSorular.Rows.Count; j++)
                        {
                            //MessageBox.Show(secilenler[i].ToString() + " " + dtKalanSorular.Rows[j]["soru_id"].ToString().Trim());
                            if (secilenler[i].ToString() == rndTumSorular.Rows[j]["soru_id"].ToString().Trim())
                            {
                                rndTumSorular.Rows[j]["soru_id"] = 0;
                                break;
                            }
                        }
                    }



                    //Seçme kontrolü
                    for (int i = 0; i < rndSecilenSorular.Rows.Count; i++)
                    {
                        //MessageBox.Show(secilenler[i].ToString());
                    }
                    //MessageBox.Show("Seçilenler bitti");

                    /*for (int i = 0; i < rndTumSorular.Rows.Count; i++)
                    {
                        MessageBox.Show(rndTumSorular.Rows[i]["soru_id"].ToString());
                    }*/

                    //rndKalanSorular.Rows.Clear();
                    for (int i = 0; i < rndTumSorular.Rows.Count; i++)
                    {
                        //MessageBox.Show(rndTumSorular.Rows[i]["soru_id"].ToString());
                        if (Convert.ToInt32(rndTumSorular.Rows[i]["soru_id"]) != 0)
                        {
                            //MessageBox.Show(rndTumSorular.Rows[i]["soru_id"].ToString());
                            rndKalanSorular.Rows.Add(rndTumSorular.Rows[i]["soru_id"].ToString().Trim(), rndTumSorular.Rows[i]["soru"].ToString().Trim(), rndTumSorular.Rows[i]["a_cevap"].ToString().Trim(), rndTumSorular.Rows[i]["b_cevap"].ToString().Trim(), rndTumSorular.Rows[i]["c_cevap"].ToString().Trim(), rndTumSorular.Rows[i]["d_cevap"].ToString().Trim(), rndTumSorular.Rows[i]["e_cevap"].ToString().Trim(), rndTumSorular.Rows[i]["dogru_cevap"].ToString().Trim(), rndTumSorular.Rows[i]["zorluk_derecesi"].ToString().Trim(), rndTumSorular.Rows[i]["soru_resim"].ToString().Trim());
                        }
                    }


                    for (int i = 0; i < rndKalanSorular.Rows.Count; i++)
                    {
                        lstRndKalan.Items.Add(rndKalanSorular.Rows[i]["soru_id"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["soru"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["a_cevap"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["b_cevap"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["c_cevap"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["d_cevap"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["e_cevap"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                        lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["soru_resim"].ToString().Trim());
                    }


                    //MessageBox.Show("Seçilmeyen soru sayısı:" + rndKalanSorular.Rows.Count.ToString());
                    //MessageBox.Show("Seçilen soru sayısı:" + rndSecilenSorular.Rows.Count.ToString());

                    lblRndToplam.Text = toplamSayi.ToString();
                    lblRndKolay.Text = kolaySayi.ToString();
                    lblRndOrta.Text = ortaSayi.ToString();
                    lblRndZor.Text = zorSayi.ToString();

                }
                else
                {
                    MessageBox.Show("Veritabanında istediğiniz kadar soru yok. Mevcut kolay seviye soru sayısı=" + totalKolay.ToString() + " .Mevcut Orta seviye soru sayısı=" + totalOrta.ToString() + " .Mevcut zor soru sayısı=" + totalZor.ToString());
                }
            }

            else
            {
                MessageBox.Show("Lütfen yıldızlı(*) alanları boş bırakmayınız.");
            }


        }

        private void btnSinavOlusturPanel_Click(object sender, EventArgs e)
        {
            this.AutoScrollPosition = new Point(0, 0);

            pcbRastgeleBaslik.Width = ekranGenislik - 160;
            pcbRastgeleBaslik.Height = 50;
            pcbRastgeleBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pcbRastgeleBaslik.Left = 0;
            pcbRastgeleBaslik.Top = 0;
            lblRastgeleBaslik.Left = 10;
            lblRastgeleBaslik.Top = 10;
            lblRastgeleBaslik.BringToFront();
            lblRastgeleBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            pnlRastgeleOlustur.Visible = true;
            pnlRastgeleOlustur.Top = 0;
            pnlRastgeleOlustur.Left = 160;
            pnlRastgeleOlustur.Width = ekranGenislik - 160;
            pnlRastgeleOlustur.Height = 800;
            pnlDersEkle.Visible = false;
            pnlSoruEkle.Visible = false;
            pnlOzelSinav.Visible = false;

            pnlRastgeleOlustur.BringToFront();

            btnDersEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSoruEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSinavOlusturPanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            button1.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSonuc.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSinavOlusturPanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            btnDersEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSoruEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            button1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSonuc.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
        }

        private void btnSoruEklePanel_Click(object sender, EventArgs e)
        {
            this.AutoScrollPosition = new Point(0, 0);

            pcbSoruEkleBaslik.Width = ekranGenislik - 160;
            pcbSoruEkleBaslik.Height = 50;
            pcbSoruEkleBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            lblSoruEkleBaslik.Left = 10;
            lblSoruEkleBaslik.Top = 10;
            lblSoruEkleBaslik.BringToFront();
            lblSoruEkleBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            pnlSoruEkle.Visible = true;
            pnlSoruEkle.Top = 0;
            pnlSoruEkle.Left = 160;
            pnlSoruEkle.Width = ekranGenislik - 160;
            pnlSoruEkle.Height = 660;
            pnlDersEkle.Visible = false;
            pnlRastgeleOlustur.Visible = false;
            pnlOzelSinav.Visible = false;

            pnlSoruEkle.BringToFront();

            btnDersEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSoruEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSinavOlusturPanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            button1.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSonuc.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSoruEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            btnDersEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSinavOlusturPanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            button1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSonuc.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
        }

        private void btnDersEklePanel_Click(object sender, EventArgs e)
        {
            this.AutoScrollPosition = new Point(0, 0);
            lblUyari.Text = "Dikkat! Dersi sildiğinizde ders\nile ilgili tüm veriler kaybolacaktır.";

            pcbDersYonBaslik.Width = ekranGenislik - 160;
            pcbDersYonBaslik.Height = 50;
            pcbDersYonBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pcbDersYonBaslik.Left = 0;
            pcbDersYonBaslik.Top = 0;
            lblDersYonBaslik.Left = 10;
            lblDersYonBaslik.Top = 10;
            lblDersYonBaslik.BringToFront();
            lblDersYonBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            pnlDersEkle.Visible = true;
            pnlDersEkle.Top = 0;
            pnlDersEkle.Left = 160;
            pnlDersEkle.Width = ekranGenislik - 160;
            pnlDersEkle.Height = 800;
            pnlSoruEkle.Visible = false;
            pnlRastgeleOlustur.Visible = false;
            pnlOzelSinav.Visible = false;

            pnlDersEkle.BringToFront();


            btnDersEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSoruEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSinavOlusturPanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            button1.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSonuc.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnDersEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            btnSoruEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSinavOlusturPanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            button1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSonuc.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
        }

        private void pnlSoruEkle_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnOzelGetir_Click(object sender, EventArgs e)
        {
            if (cmbOzelDersler.SelectedItem == null || cmbOzelZorluk.SelectedItem == null)
            {
                MessageBox.Show("Lütfen yıldızlı(*) alanları boş bırakmayınız.");
            }
            else
            {
                dataGelenSorular.Rows.Clear();
                string seviyeOzel, dersAdiOzel;
                seviyeOzel = cmbOzelZorluk.SelectedItem.ToString();

                int dersIdOzel;
                dersAdiOzel = cmbOzelDersler.SelectedItem.ToString();
                SqlCommand cmdDersID = new SqlCommand();
                cmdDersID.Connection = cnn;
                cmdDersID.CommandText = "SELECT ders_id FROM Dersler WHERE ders_adi = '" + dersAdiOzel + "'";
                DataTable dataDersID = new DataTable();
                SqlDataAdapter adpDersID = new SqlDataAdapter(cmdDersID);
                adpDersID.Fill(dataDersID);
                dersIdOzel = Convert.ToInt32(dataDersID.Rows[0]["ders_id"]);


                SqlCommand cmdSorular = new SqlCommand();
                cmdSorular.Connection = cnn;
                if (seviyeOzel == "Tümü")
                {
                    cmdSorular.CommandText = "SELECT * FROM Sorular WHERE ders_id='" + dersIdOzel + "'";
                }
                else
                {
                    cmdSorular.CommandText = "SELECT * FROM Sorular WHERE ders_id='" + dersIdOzel + "' AND zorluk_derecesi='" + seviyeOzel + "'";
                }
                SqlDataAdapter adpGelenSorular = new SqlDataAdapter(cmdSorular);
                adpGelenSorular.Fill(dataGelenSorular);

                lstGelenSoruar.Items.Clear();

                for (int i = 0; i < dataGelenSorular.Rows.Count; i++)
                {
                    lstGelenSoruar.Items.Add(dataGelenSorular.Rows[i]["soru_id"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["soru"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["a_cevap"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["b_cevap"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["c_cevap"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["d_cevap"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["e_cevap"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                    lstGelenSoruar.Items[i].SubItems.Add(dataGelenSorular.Rows[i]["soru_resim"].ToString().Trim());
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.AutoScrollPosition = new Point(0, 0);

            pcbOzelBaslik.Width = ekranGenislik - 160;
            pcbOzelBaslik.Height = 50;
            pcbOzelBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            pcbOzelBaslik.Left = 0;
            pcbOzelBaslik.Top = 0;
            lblOzelBaslik.Left = 10;
            lblOzelBaslik.Top = 10;
            lblOzelBaslik.BringToFront();
            lblOzelBaslik.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            pnlOzelSinav.Visible = true;
            pnlOzelSinav.Top = 0;
            pnlOzelSinav.Left = 160;
            pnlOzelSinav.Width = ekranGenislik - 160;
            pnlOzelSinav.Height = 800;
            pnlDersEkle.Visible = false;
            pnlSoruEkle.Visible = false;
            pnlRastgeleOlustur.Visible = false;

            pnlOzelSinav.BringToFront();

            btnDersEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSoruEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSinavOlusturPanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            button1.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSonuc.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            button1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            btnDersEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSoruEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSinavOlusturPanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSonuc.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
        }

        private void lstGelenSoruar_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstGelenSoruar.SelectedItems.Count > 0)
            {
                ListViewItem item = lstGelenSoruar.SelectedItems[0];
                seciliOzelID = Convert.ToInt32(item.SubItems[0].Text);
                if (item.SubItems[9].Text.ToString() == "" || item.SubItems[9].Text.ToString() == null)
                {
                    pcbOzelEkle.ImageLocation = @"resim-yok.jpg";
                }
                else
                {
                    pcbOzelEkle.ImageLocation = item.SubItems[9].Text.ToString();
                }
                if (item.SubItems[2].Text.ToString().Length >= 3 && item.SubItems[2].Text.ToString().Trim().Substring(0, 3) == @"C:\")
                {
                    label62.Visible = true;
                    label63.Visible = true;
                    label64.Visible = true;
                    label65.Visible = true;
                    label78.Visible = true;
                    pcbOzlA.Visible = true;
                    pcbOzlB.Visible = true;
                    pcbOzlC.Visible = true;
                    pcbOzlD.Visible = true;
                    pcbOzlE.Visible = true;
                    pcbOzlA.ImageLocation = item.SubItems[2].Text.ToString().Trim();
                    pcbOzlB.ImageLocation = item.SubItems[3].Text.ToString().Trim();
                    pcbOzlC.ImageLocation = item.SubItems[4].Text.ToString().Trim();
                    pcbOzlD.ImageLocation = item.SubItems[5].Text.ToString().Trim();
                    pcbOzlE.ImageLocation = item.SubItems[6].Text.ToString().Trim();
                }
                else
                {
                    label62.Visible = false;
                    label63.Visible = false;
                    label64.Visible = false;
                    label65.Visible = false;
                    label78.Visible = false;
                    pcbOzlA.Visible = false;
                    pcbOzlB.Visible = false;
                    pcbOzlC.Visible = false;
                    pcbOzlD.Visible = false;
                    pcbOzlE.Visible = false;
                }
                string soru;
                soru = "Zorluk Derecesi = " + item.SubItems[7].Text.ToString().Trim() + "\n" + "Soru = " + item.SubItems[1].Text.ToString().Trim() + "\n" + "a)" + item.SubItems[2].Text.ToString().Trim() + "\n" + "b)" + item.SubItems[3].Text.ToString().Trim() + "\n" + "c)" + item.SubItems[4].Text.ToString().Trim() + "\n" + "d)" + item.SubItems[5].Text.ToString().Trim() + "\n" + "e)" + item.SubItems[6].Text.ToString().Trim() + "\n" + "Doğru Cevap= " + item.SubItems[6].Text.ToString().Trim();
                rtxtOzelEkle.Text = soru;
            }
        }

        private void btnOzelEkle_Click(object sender, EventArgs e)
        {
            if (seciliOzelID.ToString() != null && seciliOzelID != 0)
            {
                secilenKolaySayi = 0;
                secilenOrtaSayi = 0;
                secilenZorSayi = 0;
                SqlCommand cmdSecilenSoru = new SqlCommand();
                cmdSecilenSoru.Connection = cnn;
                cmdSecilenSoru.CommandText = "SELECT * FROM Sorular WHERE soru_id='" + seciliOzelID + "'";
                DataTable dtSecilenSoru = new DataTable();
                SqlDataAdapter adpSecilenSoru = new SqlDataAdapter(cmdSecilenSoru);
                adpSecilenSoru.Fill(dtSecilenSoru);
                dtSecilenSorular.Rows.Add(dtSecilenSoru.Rows[0]["soru_id"], dtSecilenSoru.Rows[0]["soru"].ToString().Trim(), dtSecilenSoru.Rows[0]["a_cevap"].ToString().Trim(), dtSecilenSoru.Rows[0]["b_cevap"].ToString().Trim(), dtSecilenSoru.Rows[0]["c_cevap"].ToString().Trim(), dtSecilenSoru.Rows[0]["d_cevap"].ToString().Trim(), dtSecilenSoru.Rows[0]["e_cevap"].ToString().Trim(), dtSecilenSoru.Rows[0]["dogru_cevap"].ToString().Trim(), dtSecilenSoru.Rows[0]["zorluk_derecesi"].ToString().Trim(), dtSecilenSoru.Rows[0]["soru_resim"].ToString().Trim());



                lstOzelSinav.Items.Clear();

                for (int i = 0; i < dtSecilenSorular.Rows.Count; i++)
                {
                    /*lstOzelSinav.Items.Add(dtSecilenSorular.Rows[i]["soru_id"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["soru"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["a_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["b_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["c_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["d_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["e_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["soru_resim"].ToString().Trim());*/
                    lstOzelSinav.Items.Add(dtSecilenSorular.Rows[i]["soru_id"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["soru"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["a_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["b_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["c_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["d_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["e_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                    lstOzelSinav.Items[i].SubItems.Add(dtSecilenSorular.Rows[i]["soru_resim"].ToString().Trim());
                }

                //MessageBox.Show("Seçilen soru sorulacaklar arasına eklendi");
            }
            else
            {
                MessageBox.Show("Lütfen seçim yapınız.");
            }


            /*for (int i = 0; i < dataGelenSorular.Rows.Count; i++)
            {
                if (Convert.ToInt32(dataGelenSorular.Rows[i]["soru_id"]) == seciliOzelID)
                {
                    dataGelenSorular.Rows[i].Delete();
                    break;
                }
            }*/

            //listviewden seçileni silme
            /*for (int i = 0; i < lstGelenSoruar.Items.Count; i++)
            {
                //listView1.items[0].SubItems[0].Text;
                if (lstGelenSoruar.Items[i].SubItems[0].Text.ToString() == seciliOzelID.ToString())
                {
                    lstGelenSoruar.Items[i].Remove();
                }
            }*/
            for (int i = 0; i < dtSecilenSorular.Rows.Count; i++)
            {
                if (dtSecilenSorular.Rows[i]["zorluk_derecesi"].ToString() == "Kolay")
                {
                    secilenKolaySayi++;
                }
                else if (dtSecilenSorular.Rows[i]["zorluk_derecesi"].ToString() == "Orta")
                {
                    secilenOrtaSayi++;
                }
                else if (dtSecilenSorular.Rows[i]["zorluk_derecesi"].ToString() == "Zor")
                {
                    secilenZorSayi++;
                }
            }
            lblKolay.Text = secilenKolaySayi.ToString();
            lblOrta.Text = secilenOrtaSayi.ToString();
            lblZor.Text = secilenZorSayi.ToString();
            secilenToplam = secilenKolaySayi + secilenOrtaSayi + secilenZorSayi;
            lblToplam.Text = secilenToplam.ToString();
            pcbOzelEkle.Image = null;
            rtxtOzelEkle.Text = null;
            seciliOzelID = 0;
            label62.Visible = false;
            label63.Visible = false;
            label64.Visible = false;
            label65.Visible = false;
            label78.Visible = false;
            pcbOzlA.Visible = false;
            pcbOzlB.Visible = false;
            pcbOzlC.Visible = false;
            pcbOzlD.Visible = false;
            pcbOzlE.Visible = false;

        }

        private void lstOzelSinav_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstOzelSinav.SelectedItems.Count > 0)
            {
                ListViewItem item = lstOzelSinav.SelectedItems[0];
                seciliOzelID2 = Convert.ToInt32(item.SubItems[0].Text);
                if (item.SubItems[9].Text.ToString() == "" || item.SubItems[9].Text.ToString() == null)
                {
                    pcbOzelEkle.ImageLocation = @"resim-yok.jpg";
                }
                else
                {
                    pcbOzelEkle.ImageLocation = item.SubItems[9].Text.ToString();
                }
                if (item.SubItems[2].Text.ToString().Length >= 3 && item.SubItems[2].Text.ToString().Trim().Substring(0, 3) == @"C:\")
                {
                    label62.Visible = true;
                    label63.Visible = true;
                    label64.Visible = true;
                    label65.Visible = true;
                    label78.Visible = true;
                    pcbOzlA.Visible = true;
                    pcbOzlB.Visible = true;
                    pcbOzlC.Visible = true;
                    pcbOzlD.Visible = true;
                    pcbOzlE.Visible = true;
                    pcbOzlA.ImageLocation = item.SubItems[2].Text.ToString().Trim();
                    pcbOzlB.ImageLocation = item.SubItems[3].Text.ToString().Trim();
                    pcbOzlC.ImageLocation = item.SubItems[4].Text.ToString().Trim();
                    pcbOzlD.ImageLocation = item.SubItems[5].Text.ToString().Trim();
                    pcbOzlE.ImageLocation = item.SubItems[5].Text.ToString().Trim();
                }
                else
                {
                    label62.Visible = false;
                    label63.Visible = false;
                    label64.Visible = false;
                    label65.Visible = false;
                    label78.Visible = false;
                    pcbOzlA.Visible = false;
                    pcbOzlB.Visible = false;
                    pcbOzlC.Visible = false;
                    pcbOzlD.Visible = false;
                    pcbOzlE.Visible = false;
                }
                string soru;
                soru = "Zorluk Derecesi = " + item.SubItems[8].Text.ToString().Trim() + "\n" + "Soru = " + item.SubItems[1].Text.ToString().Trim() + "\n" + "a)" + item.SubItems[2].Text.ToString().Trim() + "\n" + "b)" + item.SubItems[3].Text.ToString().Trim() + "\n" + "c)" + item.SubItems[4].Text.ToString().Trim() + "\n" + "d)" + item.SubItems[5].Text.ToString().Trim() + "\n" + "e)" + item.SubItems[6].Text.ToString().Trim() + "\n" + "Doğru Cevap= " + item.SubItems[7].Text.ToString().Trim();
                rtxtOzelEkle.Text = soru;
            }
        }

        private void btnOzelCikar_Click(object sender, EventArgs e)
        {
            if (seciliOzelID2.ToString() != null && seciliOzelID2 != 0)
            {
                secilenKolaySayi = 0;
                secilenOrtaSayi = 0;
                secilenZorSayi = 0;
                for (int i = 0; i < lstOzelSinav.Items.Count; i++)
                {
                    //listView1.items[0].SubItems[0].Text;
                    if (lstOzelSinav.Items[i].SubItems[0].Text.ToString() == seciliOzelID2.ToString())
                    {
                        lstOzelSinav.Items[i].Remove();
                    }
                }
                for (int i = 0; i < dtSecilenSorular.Rows.Count; i++)
                {
                    if (dtSecilenSorular.Rows[i]["soru_id"].ToString() == seciliOzelID2.ToString())
                    {
                        dtSecilenSorular.Rows[i].Delete();
                    }
                }

                for (int i = 0; i < dtSecilenSorular.Rows.Count; i++)
                {
                    if (dtSecilenSorular.Rows[i]["zorluk_derecesi"].ToString() == "Kolay")
                    {
                        secilenKolaySayi++;
                    }
                    else if (dtSecilenSorular.Rows[i]["zorluk_derecesi"].ToString() == "Orta")
                    {
                        secilenOrtaSayi++;
                    }
                    else if (dtSecilenSorular.Rows[i]["zorluk_derecesi"].ToString() == "Zor")
                    {
                        secilenZorSayi++;
                    }
                }
                lblKolay.Text = secilenKolaySayi.ToString();
                lblOrta.Text = secilenOrtaSayi.ToString();
                lblZor.Text = secilenZorSayi.ToString();
                secilenToplam = secilenKolaySayi + secilenOrtaSayi + secilenZorSayi;
                lblToplam.Text = secilenToplam.ToString();
                pcbOzelEkle.Image = null;
                rtxtOzelEkle.Text = null;
                seciliOzelID2 = 0;
                //MessageBox.Show("Seçilen soru sorulacaklar arasından çıkarıldı");
            }
            else
            {
                MessageBox.Show("Lütfen seçim yapınız.");
            }
            label62.Visible = false;
            label63.Visible = false;
            label64.Visible = false;
            label65.Visible = false;
            label78.Visible = false;
            pcbOzlA.Visible = false;
            pcbOzlB.Visible = false;
            pcbOzlC.Visible = false;
            pcbOzlD.Visible = false;
            pcbOzlE.Visible = false;
        }

        private void btnDersiSil_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Ders ile ilgili tüm verileri silmek istediğinizden emin misiniz?", "Ders silme", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                int silinecekDersIdsi;
                if (cmbSilinecekDers.SelectedItem != null)
                {
                    SqlCommand cmdDersIdsi = new SqlCommand();
                    cmdDersIdsi.Connection = cnn;
                    cmdDersIdsi.CommandText = "SELECT ders_id FROM Dersler WHERE ders_adi='" + cmbSilinecekDers.SelectedItem.ToString() + "'";
                    DataTable dtDersIdsi = new DataTable();
                    SqlDataAdapter adpDersIdsi = new SqlDataAdapter(cmdDersIdsi);
                    adpDersIdsi.Fill(dtDersIdsi);
                    silinecekDersIdsi = Convert.ToInt32(dtDersIdsi.Rows[0]["ders_id"]);
                    try
                    {
                        SqlCommand cmdDerSil = new SqlCommand();
                        cmdDersIdsi.Connection = cnn;
                        cmdDersIdsi.CommandText = "DELETE FROM Dersler WHERE ders_id='" + silinecekDersIdsi + "'";
                        cmdDersIdsi.ExecuteNonQuery();

                        SqlCommand cmdDersSoruSil = new SqlCommand();
                        cmdDersSoruSil.Connection = cnn;
                        cmdDersSoruSil.CommandText = "DELETE FROM Sorular WHERE ders_id='" + silinecekDersIdsi + "'";
                        cmdDersSoruSil.ExecuteNonQuery();

                        if (Directory.Exists(@"C:\DersResimleri\dersler\" + cmbSilinecekDers.SelectedItem.ToString()))
                        {
                            FileInfo fileInfo;
                            string uzanti = ".jpg";

                            foreach (string dosya in Directory.GetFiles(@"C:\DersResimleri\dersler\" + cmbSilinecekDers.SelectedItem.ToString()))
                            {
                                fileInfo = new FileInfo(dosya);
                                if (fileInfo.Extension == uzanti) // Dosya Uzantısı ve sizin uzantınız aynımı ??
                                {
                                    fileInfo.Delete();
                                }
                            }
                            Directory.Delete(@"C:\DersResimleri\dersler\" + cmbSilinecekDers.SelectedItem.ToString());
                            //MessageBox.Show("Klasör mevcut ve siliniyor");
                        }
                        else
                        {

                        }

                        MessageBox.Show("Veriler Silindi.");

                        //Dersler ilgili comboboxlara yeniden listeleniyor
                        SqlCommand cmdDersleriGetir = new SqlCommand();
                        cmdDersleriGetir.Connection = cnn;
                        cmdDersleriGetir.CommandText = "SELECT * FROM Dersler";
                        DataTable dataDersler = new DataTable();
                        SqlDataAdapter adpDersler = new SqlDataAdapter(cmdDersleriGetir);
                        adpDersler.Fill(dataDersler);
                        cmbDersAdi.Items.Clear();
                        cmbSinavDers.Items.Clear();
                        cmbOzelDersler.Items.Clear();
                        cmbSilinecekDers.Items.Clear();
                        for (int i = 0; i < dataDersler.Rows.Count; i++)
                        {
                            cmbDersAdi.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                            cmbSinavDers.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                            cmbOzelDersler.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                            cmbSilinecekDers.Items.Add(dataDersler.Rows[i]["ders_adi"].ToString().Trim());
                        }
                        cmbDersAdi.SelectedItem = null;
                        cmbSinavDers.SelectedItem = null;
                        cmbOzelDersler.SelectedItem = null;
                        cmbSilinecekDers.SelectedItem = null;
                    }
                    catch (Exception hata)
                    {
                        MessageBox.Show("Veriler silinemedi. Hata mesajı= " + hata.Message);
                    }

                }
                else
                {
                    MessageBox.Show("Lütfen yıldızlı(*) alanları boş bırakmayınız.");
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                MessageBox.Show("Ders silinmedi.");
            }

        }

        private void btnOnizleme_Click(object sender, EventArgs e)
        {
            string oDersAdi, oSoru, oA = "", oB = "", oC = "", oD = "", oE = "", oDogru, oZorluk, oResim, yazilacak;
            if ((rdbA.Checked == true || rdbB.Checked == true || rdbC.Checked == true || rdbD.Checked == true || rdbE.Checked == true) && rtxtSoru.Text != "" &&/* txtA.Text != "" && txtB.Text != "" && txtC.Text != "" && txtD.Text != "" &&*/ cmbDersAdi.SelectedItem != null && cmbZorluk.SelectedItem != null)
            {
                oDersAdi = cmbDersAdi.SelectedItem.ToString();
                oSoru = rtxtSoru.Text;
                if (rdbSikMetin.Checked == true)
                {
                    oA = txtA.Text;
                    oB = txtB.Text;
                    oC = txtC.Text;
                    oD = txtD.Text;
                    oE = txtE.Text;
                }
                else if (rdbSikResim.Checked == true)
                {
                    oA = aResim;
                    oB = bResim;
                    oC = cResim;
                    oD = dResim;
                    oE = eResim;
                    pcbA.ImageLocation = aResim;
                    pcbB.ImageLocation = bResim;
                    pcbC.ImageLocation = cResim;
                    pcbD.ImageLocation = dResim;
                    pcbE.ImageLocation = eResim;
                }
                if (rdbA.Checked == true)
                {
                    oDogru = oA;
                }
                else if (rdbB.Checked == true)
                {
                    oDogru = oB;
                }
                else if (rdbC.Checked == true)
                {
                    oDogru = oC;
                }
                else if (rdbD.Checked == true)
                {
                    oDogru = oD;
                }
                else
                {
                    oDogru = oE;
                }
                oZorluk = cmbZorluk.SelectedItem.ToString();
                //MessageBox.Show(resimYolu);
                if (resimYolu == null || resimYolu == "")
                {
                    //MessageBox.Show("İfe girdi");
                    pcbOnizleme.ImageLocation = @"resim-yok.jpg";
                }
                else
                {
                    oResim = resimYolu;
                    pcbOnizleme.ImageLocation = oResim;
                }
                yazilacak = oDersAdi + "\n" + "Zorluk seviyesi =" + oZorluk + "\n" + oSoru + "\nA)" + oA + "\nB)" + oB + "\nC)" + oC + "\nD)" + oD + "\n" + "E)" + oE + "\n" + "Doğru Cevap=" + oDogru;
                rtxtOnizleme.Text = yazilacak;

            }
            else
            {
                MessageBox.Show("Lütfen yıldızlı(*) alanları boş bırakmayınız.");
            }

        }

        private void lstRndGelen_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(lstRndGelen.SelectedItems[0].SubItems[columnHeader24].Text.ToString());
            if (lstRndGelen.SelectedItems.Count > 0)
            {
                pcbRndA.Image = null;
                pcbRndB.Image = null;
                pcbRndC.Image = null;
                pcbRndD.Image = null;
                pcbRndE.Image = null;
                ListViewItem item = lstRndGelen.SelectedItems[0];
                rndSeciliId = Convert.ToInt32(item.SubItems[0].Text);
                if (item.SubItems[9].Text.ToString() == "")
                {

                    pcbRndOnizleme.ImageLocation = @"resim-yok.jpg";
                }
                else
                {
                    pcbRndOnizleme.ImageLocation = item.SubItems[9].Text.ToString();
                }

                if (item.SubItems[2].Text.ToString().Length >= 3 && item.SubItems[2].Text.ToString().Trim().Substring(0, 3) == @"C:\")
                {
                    label58.Visible = true;
                    label59.Visible = true;
                    label60.Visible = true;
                    label61.Visible = true;
                    label79.Visible = true;
                    pcbRndA.Visible = true;
                    pcbRndB.Visible = true;
                    pcbRndC.Visible = true;
                    pcbRndD.Visible = true;
                    pcbRndE.Visible = true;
                    pcbRndA.ImageLocation = item.SubItems[2].Text.ToString().Trim();
                    pcbRndB.ImageLocation = item.SubItems[3].Text.ToString().Trim();
                    pcbRndC.ImageLocation = item.SubItems[4].Text.ToString().Trim();
                    pcbRndD.ImageLocation = item.SubItems[5].Text.ToString().Trim();
                    pcbRndE.ImageLocation = item.SubItems[5].Text.ToString().Trim();
                }
                else
                {
                    label58.Visible = false;
                    label59.Visible = false;
                    label60.Visible = false;
                    label61.Visible = false;
                    label79.Visible = false;
                    pcbRndA.Visible = false;
                    pcbRndB.Visible = false;
                    pcbRndC.Visible = false;
                    pcbRndD.Visible = false;
                    pcbRndE.Visible = false;
                }
                string soru;
                soru = "Zorluk Derecesi = " + item.SubItems[8].Text.ToString().Trim() + "\n" + "Soru = " + item.SubItems[1].Text.ToString().Trim() + "\n" + "a)" + item.SubItems[2].Text.ToString().Trim() + "\n" + "b)" + item.SubItems[3].Text.ToString().Trim() + "\n" + "c)" + item.SubItems[4].Text.ToString().Trim() + "\n" + "d)" + item.SubItems[5].Text.ToString().Trim() + "\n" + "e)" + item.SubItems[6].Text.ToString().Trim() + "\n" + "Doğru Cevap= " + item.SubItems[7].Text.ToString().Trim();
                rtxtRndOnizleme.Text = soru;
            }
        }

        private void lstRndKalan_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstRndKalan.SelectedItems.Count > 0)
            {
                ListViewItem item = lstRndKalan.SelectedItems[0];
                rndSeciliIDEkle = Convert.ToInt32(item.SubItems[0].Text);
                if (item.SubItems[9].Text.ToString() == "")
                {
                    pcbRndOnizleme.ImageLocation = @"resim-yok.jpg";
                }
                else
                {
                    pcbRndOnizleme.ImageLocation = item.SubItems[9].Text.ToString();
                }

                if (item.SubItems[2].Text.ToString().Length >= 3 && item.SubItems[2].Text.ToString().Trim().Substring(0, 3) == @"C:\")
                {
                    label58.Visible = true;
                    label59.Visible = true;
                    label60.Visible = true;
                    label61.Visible = true;
                    label79.Visible = true;
                    pcbRndA.Visible = true;
                    pcbRndB.Visible = true;
                    pcbRndC.Visible = true;
                    pcbRndD.Visible = true;
                    pcbRndE.Visible = true;
                    pcbRndA.ImageLocation = item.SubItems[2].Text.ToString().Trim();
                    pcbRndB.ImageLocation = item.SubItems[3].Text.ToString().Trim();
                    pcbRndC.ImageLocation = item.SubItems[4].Text.ToString().Trim();
                    pcbRndD.ImageLocation = item.SubItems[5].Text.ToString().Trim();
                    pcbRndE.ImageLocation = item.SubItems[6].Text.ToString().Trim();
                }
                else
                {
                    label58.Visible = false;
                    label59.Visible = false;
                    label60.Visible = false;
                    label61.Visible = false;
                    label79.Visible = false;
                    pcbRndA.Visible = false;
                    pcbRndB.Visible = false;
                    pcbRndC.Visible = false;
                    pcbRndD.Visible = false;
                    pcbRndE.Visible = false;
                }
                string soru;

                soru = "Zorluk Derecesi = " + item.SubItems[8].Text.ToString().Trim() + "\n" + "Soru = " + item.SubItems[1].Text.ToString().Trim() + "\n" + "a)" + item.SubItems[2].Text.ToString().Trim() + "\n" + "b)" + item.SubItems[3].Text.ToString().Trim() + "\n" + "c)" + item.SubItems[4].Text.ToString().Trim() + "\n" + "d)" + item.SubItems[5].Text.ToString().Trim() + "\n" + "e)" + item.SubItems[6].Text.ToString().Trim() + "\n" + "Doğru Cevap= " + item.SubItems[7].Text.ToString().Trim();
                rtxtRndOnizleme.Text = soru;
            }
        }

        private void btnRndCikar_Click(object sender, EventArgs e)
        {
            string cikarilan;
            if (rndSeciliId != 0)
            {
                for (int i = 0; i < rndSecilenSorular.Rows.Count; i++)
                {
                    if (rndSecilenSorular.Rows[i]["soru_id"].ToString() == rndSeciliId.ToString())
                    {
                        cikarilan = rndSecilenSorular.Rows[i]["zorluk_derecesi"].ToString().Trim();
                        rndKalanSorular.Rows.Add(rndSecilenSorular.Rows[i]["soru_id"].ToString().Trim(), rndSecilenSorular.Rows[i]["soru"].ToString().Trim(), rndSecilenSorular.Rows[i]["a_cevap"].ToString().Trim(), rndSecilenSorular.Rows[i]["b_cevap"].ToString().Trim(), rndSecilenSorular.Rows[i]["c_cevap"].ToString().Trim(), rndSecilenSorular.Rows[i]["d_cevap"].ToString().Trim(), rndSecilenSorular.Rows[i]["e_cevap"].ToString().Trim(), rndSecilenSorular.Rows[i]["dogru_cevap"].ToString().Trim(), rndSecilenSorular.Rows[i]["zorluk_derecesi"].ToString().Trim(), rndSecilenSorular.Rows[i]["soru_resim"].ToString().Trim());
                        rndSecilenSorular.Rows[i].Delete();
                        MessageBox.Show("Seçili soru silindi.");
                        if (cikarilan == "Kolay")
                        {
                            kolaySayi--;
                        }
                        else if (cikarilan == "Orta")
                        {
                            ortaSayi--;
                        }
                        else if (cikarilan == "Zor")
                        {
                            zorSayi--;
                        }
                        toplamSayi--;
                        lblRndToplam.Text = toplamSayi.ToString();
                        lblRndKolay.Text = kolaySayi.ToString();
                        lblRndOrta.Text = ortaSayi.ToString();
                        lblRndZor.Text = zorSayi.ToString();
                        break;
                    }

                }

                lstRndGelen.Items.Clear();
                for (int i = 0; i < rndSecilenSorular.Rows.Count; i++)
                {
                    lstRndGelen.Items.Add(rndSecilenSorular.Rows[i]["soru_id"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["soru"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["a_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["b_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["c_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["d_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["e_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["soru_resim"].ToString().Trim());
                }

                pcbRndOnizleme.Image = null;
                rtxtRndOnizleme.Text = null;
                lstRndKalan.Items.Clear();
                for (int i = 0; i < rndKalanSorular.Rows.Count; i++)
                {
                    lstRndKalan.Items.Add(rndKalanSorular.Rows[i]["soru_id"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["soru"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["a_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["b_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["c_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["d_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["e_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["soru_resim"].ToString().Trim());
                }
                //MessageBox.Show(rndSecilenSorular.Rows.Count.ToString());
                //MessageBox.Show(rndKalanSorular.Rows.Count.ToString());


                label58.Visible = false;
                label59.Visible = false;
                label60.Visible = false;
                label61.Visible = false;
                label79.Visible = false;
                pcbRndA.Visible = false;
                pcbRndB.Visible = false;
                pcbRndC.Visible = false;
                pcbRndD.Visible = false;
                pcbRndE.Visible = false;
                rndSeciliId = 0;
            }
            else
            {
                MessageBox.Show("Lütfen soru seçin.");
            }

        }

        private void btnRndEkle_Click(object sender, EventArgs e)
        {
            string eklenen;
            if (rndSeciliIDEkle != 0)
            {
                for (int i = 0; i < rndKalanSorular.Rows.Count; i++)
                {
                    if (rndKalanSorular.Rows[i]["soru_id"].ToString() == rndSeciliIDEkle.ToString())
                    {
                        eklenen = rndKalanSorular.Rows[i]["zorluk_derecesi"].ToString().Trim();
                        rndSecilenSorular.Rows.Add(rndKalanSorular.Rows[i]["soru_id"].ToString().Trim(), rndKalanSorular.Rows[i]["soru"].ToString().Trim(), rndKalanSorular.Rows[i]["a_cevap"].ToString().Trim(), rndKalanSorular.Rows[i]["b_cevap"].ToString().Trim(), rndKalanSorular.Rows[i]["c_cevap"].ToString().Trim(), rndKalanSorular.Rows[i]["d_cevap"].ToString().Trim(), rndKalanSorular.Rows[i]["e_cevap"].ToString().Trim(), rndKalanSorular.Rows[i]["dogru_cevap"].ToString().Trim(), rndKalanSorular.Rows[i]["zorluk_derecesi"].ToString().Trim(), rndKalanSorular.Rows[i]["soru_resim"].ToString().Trim());
                        rndKalanSorular.Rows[i].Delete();
                        //MessageBox.Show("Seçilen soru sınav soruları arasına eklendi");
                        if (eklenen == "Kolay")
                        {
                            kolaySayi++;
                        }
                        else if (eklenen == "Orta")
                        {
                            ortaSayi++;
                        }
                        else if (eklenen == "Zor")
                        {
                            zorSayi++;
                        }
                        toplamSayi++;
                        lblRndToplam.Text = toplamSayi.ToString();
                        lblRndKolay.Text = kolaySayi.ToString();
                        lblRndOrta.Text = ortaSayi.ToString();
                        lblRndZor.Text = zorSayi.ToString();
                        break;
                    }
                }
                label58.Visible = false;
                label59.Visible = false;
                label60.Visible = false;
                label61.Visible = false;
                label79.Visible = false;
                pcbRndA.Visible = false;
                pcbRndB.Visible = false;
                pcbRndC.Visible = false;
                pcbRndD.Visible = false;
                pcbRndE.Visible = false;


                lstRndGelen.Items.Clear();
                for (int i = 0; i < rndSecilenSorular.Rows.Count; i++)
                {
                    lstRndGelen.Items.Add(rndSecilenSorular.Rows[i]["soru_id"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["soru"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["a_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["b_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["c_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["d_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["e_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                    lstRndGelen.Items[i].SubItems.Add(rndSecilenSorular.Rows[i]["soru_resim"].ToString().Trim());
                }

                lstRndKalan.Items.Clear();
                for (int i = 0; i < rndKalanSorular.Rows.Count; i++)
                {
                    lstRndKalan.Items.Add(rndKalanSorular.Rows[i]["soru_id"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["soru"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["a_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["b_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["c_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["d_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["e_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["dogru_cevap"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["zorluk_derecesi"].ToString().Trim());
                    lstRndKalan.Items[i].SubItems.Add(rndKalanSorular.Rows[i]["soru_resim"].ToString().Trim());
                }


                //MessageBox.Show(rndSecilenSorular.Rows.Count.ToString());
                //MessageBox.Show(rndKalanSorular.Rows.Count.ToString());

                pcbRndOnizleme.ImageLocation = null;
                rtxtRndOnizleme.Text = null;
                rndSeciliIDEkle = 0;
            }
            else
            {
                MessageBox.Show("Lütfen soru seçin.");
            }

        }

        private void btnRndWord_Click(object sender, EventArgs e)
        {
            /*int resimliSiklar = 0;
            int resimliSorular = 0;*/
            int soruSayisi = 0;
            /*for (int i = 0; i < rndSecilenSorular.Rows.Count; i++)
            {
                if (rndSecilenSorular.Rows[i]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                {
                    resimliSiklar++;
                }
                if (rndSecilenSorular.Rows[i]["soru_resim"].ToString().Trim() != "")
                {
                    resimliSorular++;
                }
            }*/
            if (rndSecilenSorular.Rows.Count > 0)
            {
                if ((cmbRndGrup.SelectedItem.ToString() != "" || cmbRndGrup.SelectedItem != null) && (cmbRndVF.SelectedItem.ToString() != "" || cmbRndVF.SelectedItem != null) && txtRndBolum.Text != "" && txtRndSure.Text != "" && (cmbSinavDers.SelectedItem.ToString() != "" || cmbSinavDers.SelectedItem != null))
                {
                    int tabloSatirSayisi = 0;
                    int sikSayisi = 0, resimSayisi = 0;

                    soruSayisi = rndSecilenSorular.Rows.Count;
                    sikSayisi = soruSayisi;
                    for (int i = 0; i < soruSayisi; i++)
                    {
                        if (rndSecilenSorular.Rows[i]["soru_resim"].ToString() != null && rndSecilenSorular.Rows[i]["soru_resim"].ToString() != "")
                        {
                            resimSayisi++;
                        }
                    }
                    //tabloSatirSayisi = (dtA.Rows.Count * 2) + resimliSorular + (resimliSiklar);
                    tabloSatirSayisi = soruSayisi + sikSayisi + resimSayisi;
                    //MessageBox.Show("toplam satır= "+tabloSatirSayisi.ToString()+ " toplam Soru= "+soruSayisi.ToString()+" Resim sayısı="+resimSayisi.ToString());

                    /*yazdirilacakTur = "Rastgele";
                    Form4 frm4 = new Form4();
                    frm4.Show();*/
                    if (cmbRndGrup.SelectedItem.ToString() == "Tek Grup")
                    {
                        object omissing = System.Reflection.Missing.Value;
                        object son = "\\endofdoc";

                        word.Application olustur;
                        word.Document icerik;
                        olustur = new word.Application();
                        olustur.Visible = true;
                        icerik = olustur.Documents.Add(ref omissing);

                        icerik.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme

                        //icerik.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = richTextBox1.Text;

                        /*word.Range headerRange = icerik.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Fields.Add(headerRange, word.WdFieldType.wdFieldPage);
                        headerRange.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        //headerRange.Font.ColorIndex = word.WdColorIndex.wdRed;
                        headerRange.Font.Size = 10;
                        headerRange.Font.Bold = 1;
                        //headerRange.Font.Name = "Arial";
                        headerRange.Text = rtxtRndUstbilgi.Text + "\n\nAd:                                            Soyad:                                            No:                                            \n";
                        */

                        word.Range headerRangeA = icerik.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeA.Fields.Add(headerRangeA, word.WdFieldType.wdFieldPage);
                        headerRangeA.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeA.Font.Size = 10;
                        headerRangeA.Font.Bold = 1;
                        string sinavKontrol = cmbRndVF.SelectedItem.ToString();
                        if (sinavKontrol == "Vize")
                        {
                            sinavKontrol = "VİZESİ";
                        }
                        else if (sinavKontrol == "Final")
                        {
                            sinavKontrol = "FİNALİ";
                        }
                        string ustbilgi = "MARMARA ÜNİVERSİTESİ TEKNİK BİLİMLER MESLEK YÜKSEKOKULU\n" + txtRndBolum.Text.ToString().ToUpper() + " PROGRAMI " + cmbSinavDers.SelectedItem.ToString().ToUpper() + " DERSİ " + sinavKontrol.ToString() + "\nSINAV SÜRESİ=" + txtRndSure.Text;
                        headerRangeA.Text = ustbilgi + "\n\nAd:                                            Soyad:                                            No:                                            \n";




                        word.Table oTable;
                        word.Range wrdRng = icerik.Bookmarks.get_Item(ref son).Range;
                        oTable = icerik.Tables.Add(wrdRng, tabloSatirSayisi, 1, ref omissing, ref omissing);
                        oTable.Range.ParagraphFormat.SpaceAfter = 6;
                        oTable.Range.Font.Size = 10;

                        /*int[] soruNo = new int [rndSecilenSorular.Rows.Count];
                        string[] soruCumlesi = new string[rndSecilenSorular.Rows.Count];*/

                        string soru = "";
                        string a = "";
                        string b = "";
                        string c = "";
                        string d = "";
                        string eC = "";
                        //string dogru = "";
                        int sayac = 0;
                        string siklar = "";
                        string tumSoru = "";
                        string soruResmi = "";

                        for (int i = 1; i <= tabloSatirSayisi; i++)
                        {
                            //MessageBox.Show("sayaç= "+sayac.ToString()+ " i= "+i.ToString());
                            soru = rndSecilenSorular.Rows[sayac]["soru"].ToString().Trim();
                            a = rndSecilenSorular.Rows[sayac]["a_cevap"].ToString().Trim();
                            b = rndSecilenSorular.Rows[sayac]["b_cevap"].ToString().Trim();
                            c = rndSecilenSorular.Rows[sayac]["c_cevap"].ToString().Trim();
                            d = rndSecilenSorular.Rows[sayac]["d_cevap"].ToString().Trim();
                            eC = rndSecilenSorular.Rows[sayac]["e_cevap"].ToString().Trim();
                            siklar = "a)" + a + "\nb)" + b + "\nc)" + c + "\nd)" + d + "\ne)" + eC + "\n";
                            tumSoru = (sayac + 1).ToString() + "-)" + soru;
                            soruResmi = rndSecilenSorular.Rows[sayac]["soru_resim"].ToString().Trim();
                            if (soruResmi != null && soruResmi != "")
                            {

                                //MessageBox.Show(soruResmi);
                                oTable.Cell(i, 1).Range.InlineShapes.AddPicture((@soruResmi), Type.Missing, Type.Missing, Type.Missing);
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 1;

                                oTable.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 0;
                                //oTable.Cell(i, 1).Range.Text = siklar;
                                if (rndSecilenSorular.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (rndSecilenSorular.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        //oTable.Rows[i].Cells[1].SplitCell(3, 2);
                                        oTable.Rows[i].Cells[1].Split(3, 4);
                                        oTable.Cell(i, 1).Range.Text = "a)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "b)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "c)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "d)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "e)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTable.Cell(i, 1).Range.Text = siklar;
                                    }
                                }
                                else
                                {
                                    oTable.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            else
                            {

                                oTable.Cell(i, 1).Range.Font.Bold = 1;

                                oTable.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 0;
                                //oTable.Cell(i, 1).Range.Text = siklar;
                                if (rndSecilenSorular.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (rndSecilenSorular.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        oTable.Rows[i].Cells[1].Split(3, 4);

                                        oTable.Cell(i, 1).Range.Text = "a)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "b)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "c)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "d)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "e)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTable.Cell(i, 1).Range.Text = siklar;
                                    }
                                }
                                else
                                {
                                    oTable.Cell(i, 1).Range.Text = siklar;
                                }

                            }
                            sayac++;

                        }


                        //------Cevap Anahtarı

                        object omissing2 = System.Reflection.Missing.Value;
                        object son2 = "\\endofdoc";

                        word.Application cevapAnahtari;
                        word.Document icerik2;
                        cevapAnahtari = new word.Application();
                        cevapAnahtari.Visible = true;
                        icerik2 = cevapAnahtari.Documents.Add(ref omissing2);

                        icerik2.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        /*word.Range headerRangeCevaplar = icerik2.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplar.Fields.Add(headerRangeCevaplar, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplar.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplar.Font.Size = 10;
                        headerRangeCevaplar.Font.Bold = 1;
                        headerRangeCevaplar.Text = rtxtRndUstbilgi.Text + "\nCEVAP ANAHTARI\n";
                        */
                        word.Range headerRangeCevaplar = icerik2.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplar.Fields.Add(headerRangeCevaplar, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplar.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplar.Font.Size = 10;
                        headerRangeCevaplar.Font.Bold = 1;
                        //string sinavKontrol = cmbRndVF.SelectedItem.ToString();
                        if (sinavKontrol == "Vize")
                        {
                            sinavKontrol = "VİZESİ";
                        }
                        else if (sinavKontrol == "Final")
                        {
                            sinavKontrol = "FİNALİ";
                        }
                        ustbilgi = "MARMARA ÜNİVERSİTESİ TEKNİK BİLİMLER MESLEK YÜKSEKOKULU\n" + txtRndBolum.Text.ToString().ToUpper() + " PROGRAMI " + cmbSinavDers.SelectedItem.ToString().ToUpper() + " DERSİ " + sinavKontrol.ToString() + "\nSINAV SÜRESİ=" + txtRndSure.Text;
                        headerRangeCevaplar.Text = ustbilgi + "\nCEVAP ANAHTARI";


                        word.Table oTableC;
                        word.Range wrdRngC = icerik2.Bookmarks.get_Item(ref son2).Range;
                        oTableC = icerik2.Tables.Add(wrdRngC, tabloSatirSayisi, 1, ref omissing2, ref omissing2);
                        oTableC.Range.ParagraphFormat.SpaceAfter = 6;
                        oTableC.Range.Font.Size = 10;

                        string soru2 = "";
                        int sayac2 = 0;
                        int cevapSatirSayisi = soruSayisi * 2;
                        string tumSoru2 = "";
                        string dogruSik = "";
                        string dogruCevap = "";
                        for (int i = 1; i <= cevapSatirSayisi; i++)
                        {
                            soru2 = rndSecilenSorular.Rows[sayac2]["soru"].ToString().Trim();
                            tumSoru2 = (sayac2 + 1).ToString() + "-)" + soru2;
                            dogruCevap = rndSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim();
                            if (rndSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == rndSecilenSorular.Rows[sayac2]["a_cevap"].ToString().Trim())
                            {
                                dogruSik = "A-)";
                            }
                            else if (rndSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == rndSecilenSorular.Rows[sayac2]["b_cevap"].ToString().Trim())
                            {
                                dogruSik = "B-)";
                            }
                            else if (rndSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == rndSecilenSorular.Rows[sayac2]["c_cevap"].ToString().Trim())
                            {
                                dogruSik = "C-)";
                            }
                            else if (rndSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == rndSecilenSorular.Rows[sayac2]["d_cevap"].ToString().Trim())
                            {
                                dogruSik = "D-)";
                            }
                            else if (rndSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == rndSecilenSorular.Rows[sayac2]["e_cevap"].ToString().Trim())
                            {
                                dogruSik = "E-)";
                            }
                            oTableC.Cell(i, 1).Range.Font.Bold = 1;
                            oTableC.Cell(i, 1).Range.Text = tumSoru2;
                            i++;
                            oTableC.Cell(i, 1).Range.Font.Bold = 0;
                            oTableC.Cell(i, 1).Range.Text = dogruSik + dogruCevap;
                            sayac2++;

                        }
                    }
                    if (cmbRndGrup.SelectedItem.ToString() == "İki Grup")
                    {
                        DataTable dtA = new DataTable();
                        DataTable dtB = new DataTable();

                        dtA.Columns.Add("soru_id");
                        dtA.Columns.Add("soru");
                        dtA.Columns.Add("a_cevap");
                        dtA.Columns.Add("b_cevap");
                        dtA.Columns.Add("c_cevap");
                        dtA.Columns.Add("d_cevap");
                        dtA.Columns.Add("dogru_cevap");
                        dtA.Columns.Add("zorluk_derecesi");
                        dtA.Columns.Add("soru_resim");

                        dtB.Columns.Add("soru_id");
                        dtB.Columns.Add("soru");
                        dtB.Columns.Add("a_cevap");
                        dtB.Columns.Add("b_cevap");
                        dtB.Columns.Add("c_cevap");
                        dtB.Columns.Add("d_cevap");
                        dtB.Columns.Add("dogru_cevap");
                        dtB.Columns.Add("zorluk_derecesi");
                        dtB.Columns.Add("soru_resim");

                        //dtA.DefaultView.Sort = "soru_id DESC";
                        dtA.Merge(rndSecilenSorular);

                        dtB.Merge(rndSecilenSorular);
                        DataView dw = dtB.DefaultView;
                        dw.Sort = "a_cevap";
                        dtB = dw.ToTable();


                        object omissing = System.Reflection.Missing.Value;
                        object son = "\\endofdoc";

                        word.Application olustur;
                        word.Document icerik;
                        olustur = new word.Application();
                        olustur.Visible = true;
                        icerik = olustur.Documents.Add(ref omissing);

                        icerik.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        /*word.Range headerRange = icerik.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Fields.Add(headerRange, word.WdFieldType.wdFieldPage);
                        headerRange.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRange.Font.Size = 10;
                        headerRange.Font.Bold = 1;
                        headerRange.Text = rtxtRndUstbilgi.Text + "\n\nAd:                                            Soyad:                                            No:                                            \nA GRUBU \n";
                        */
                        word.Range headerRangeA = icerik.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeA.Fields.Add(headerRangeA, word.WdFieldType.wdFieldPage);
                        headerRangeA.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeA.Font.Size = 10;
                        headerRangeA.Font.Bold = 1;
                        string sinavKontrol = cmbRndVF.SelectedItem.ToString();
                        if (sinavKontrol == "Vize")
                        {
                            sinavKontrol = "VİZESİ";
                        }
                        else if (sinavKontrol == "Final")
                        {
                            sinavKontrol = "FİNALİ";
                        }
                        string ustbilgi = "MARMARA ÜNİVERSİTESİ TEKNİK BİLİMLER MESLEK YÜKSEKOKULU\n" + txtRndBolum.Text.ToString().ToUpper() + " PROGRAMI " + cmbSinavDers.SelectedItem.ToString().ToUpper() + " DERSİ " + sinavKontrol.ToString() + "\nSINAV SÜRESİ=" + txtRndSure.Text.ToString().ToUpper();
                        headerRangeA.Text = ustbilgi + "\n\nAd:                                            Soyad:                                            No:                                            \nA GURUBU";


                        word.Table oTable;
                        word.Range wrdRng = icerik.Bookmarks.get_Item(ref son).Range;
                        oTable = icerik.Tables.Add(wrdRng, tabloSatirSayisi, 1, ref omissing, ref omissing);
                        oTable.Range.ParagraphFormat.SpaceAfter = 6;
                        oTable.Range.Font.Size = 10;


                        string soru = "";
                        string a = "";
                        string b = "";
                        string c = "";
                        string d = "";
                        string eC = "";
                        //string dogru = "";
                        int sayac = 0;
                        string siklar = "";
                        string tumSoru = "";
                        string soruResmi = "";

                        for (int i = 1; i <= tabloSatirSayisi; i++)
                        {
                            soru = dtA.Rows[sayac]["soru"].ToString().Trim();
                            a = dtA.Rows[sayac]["a_cevap"].ToString().Trim();
                            b = dtA.Rows[sayac]["b_cevap"].ToString().Trim();
                            c = dtA.Rows[sayac]["c_cevap"].ToString().Trim();
                            d = dtA.Rows[sayac]["d_cevap"].ToString().Trim();
                            eC = dtA.Rows[sayac]["e_cevap"].ToString().Trim();
                            siklar = "a)" + a + "\nb)" + b + "\nc)" + c + "\nd)" + d + "\ne)" + eC + "\n";
                            tumSoru = (sayac + 1).ToString() + "-)" + soru;
                            soruResmi = dtA.Rows[sayac]["soru_resim"].ToString().Trim();
                            if (soruResmi != null && soruResmi != "")
                            {

                                oTable.Cell(i, 1).Range.InlineShapes.AddPicture((@soruResmi), Type.Missing, Type.Missing, Type.Missing);
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 1;

                                oTable.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 0;

                                if (dtA.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtA.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        //oTable.Rows[i].Cells[1].SplitCell(3, 2);
                                        oTable.Rows[i].Cells[1].Split(3, 4);
                                        oTable.Cell(i, 1).Range.Text = "a)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "b)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "c)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "d)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "e)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTable.Cell(i, 1).Range.Text = siklar;
                                    }
                                }

                                else
                                {
                                    oTable.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            else
                            {

                                oTable.Cell(i, 1).Range.Font.Bold = 1;

                                oTable.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 0;

                                if (dtA.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtA.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        oTable.Rows[i].Cells[1].Split(3, 4);

                                        oTable.Cell(i, 1).Range.Text = "a)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "b)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "c)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "d)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "e)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTable.Cell(i, 1).Range.Text = siklar;
                                    }
                                }
                                else
                                {
                                    oTable.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            sayac++;

                        }


                        //------A Cevap Anahtarı

                        object omissing2 = System.Reflection.Missing.Value;
                        object son2 = "\\endofdoc";

                        word.Application cevapAnahtari;
                        word.Document icerik2;
                        cevapAnahtari = new word.Application();
                        cevapAnahtari.Visible = true;
                        icerik2 = cevapAnahtari.Documents.Add(ref omissing2);

                        icerik2.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        /*word.Range headerRangeCevaplar = icerik2.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplar.Fields.Add(headerRangeCevaplar, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplar.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplar.Font.Size = 10;
                        headerRangeCevaplar.Font.Bold = 1;
                        headerRangeCevaplar.Text = rtxtRndUstbilgi.Text + "\nA GRUBU CEVAP ANAHTARI\n";
                        */
                        word.Range headerRangeCevaplar = icerik2.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplar.Fields.Add(headerRangeCevaplar, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplar.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplar.Font.Size = 10;
                        headerRangeCevaplar.Font.Bold = 1;
                        //string sinavKontrol = cmbRndVF.SelectedItem.ToString();
                        if (sinavKontrol == "Vize")
                        {
                            sinavKontrol = "VİZESİ";
                        }
                        else if (sinavKontrol == "Final")
                        {
                            sinavKontrol = "FİNALİ";
                        }
                        ustbilgi = "MARMARA ÜNİVERSİTESİ TEKNİK BİLİMLER MESLEK YÜKSEKOKULU\n" + txtRndBolum.Text.ToString().ToUpper() + " PROGRAMI " + cmbSinavDers.SelectedItem.ToString().ToUpper() + " DERSİ " + sinavKontrol.ToString() + "\nSINAV SÜRESİ=" + txtRndSure.Text.ToString().ToUpper();
                        headerRangeCevaplar.Text = ustbilgi + "\nA GURUBU CEVAP ANAHTARI";


                        word.Table oTableC;
                        word.Range wrdRngC = icerik2.Bookmarks.get_Item(ref son2).Range;
                        oTableC = icerik2.Tables.Add(wrdRngC, tabloSatirSayisi, 1, ref omissing2, ref omissing2);
                        oTableC.Range.ParagraphFormat.SpaceAfter = 6;
                        oTableC.Range.Font.Size = 10;

                        string soru2 = "";
                        int sayac2 = 0;
                        int cevapSatirSayisi = soruSayisi * 2;
                        string tumSoru2 = "";
                        string dogruSik = "";
                        string dogruCevap = "";
                        for (int i = 1; i <= cevapSatirSayisi; i++)
                        {
                            soru2 = dtA.Rows[sayac2]["soru"].ToString().Trim();
                            tumSoru2 = (sayac2 + 1).ToString() + "-)" + soru2;
                            dogruCevap = dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim();
                            if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["a_cevap"].ToString().Trim())
                            {
                                dogruSik = "A-)";
                            }
                            else if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["b_cevap"].ToString().Trim())
                            {
                                dogruSik = "B-)";
                            }
                            else if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["c_cevap"].ToString().Trim())
                            {
                                dogruSik = "C-)";
                            }
                            else if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["d_cevap"].ToString().Trim())
                            {
                                dogruSik = "D-)";
                            }
                            else if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["e_cevap"].ToString().Trim())
                            {
                                dogruSik = "E-)";
                            }
                            oTableC.Cell(i, 1).Range.Font.Bold = 1;
                            oTableC.Cell(i, 1).Range.Text = tumSoru2;
                            i++;
                            oTableC.Cell(i, 1).Range.Font.Bold = 0;
                            oTableC.Cell(i, 1).Range.Text = dogruSik + dogruCevap;
                            sayac2++;





                        }
                        // B GRUBU

                        object omissingB = System.Reflection.Missing.Value;
                        object sonB = "\\endofdoc";

                        word.Application olusturB;
                        word.Document icerikB;
                        olusturB = new word.Application();
                        olusturB.Visible = true;
                        icerikB = olusturB.Documents.Add(ref omissingB);

                        icerikB.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme
                        /*

                        word.Range headerRangeB = icerikB.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeB.Fields.Add(headerRangeB, word.WdFieldType.wdFieldPage);
                        headerRangeB.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeB.Font.Size = 10;
                        headerRangeB.Font.Bold = 1;
                        headerRangeB.Text = rtxtRndUstbilgi.Text + "\n\nAd:                                            Soyad:                                            No:                                            \nB GRUBU \n";
                        */
                        word.Range headerRangeB = icerikB.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeB.Fields.Add(headerRangeB, word.WdFieldType.wdFieldPage);
                        headerRangeB.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeB.Font.Size = 10;
                        headerRangeB.Font.Bold = 1;
                        sinavKontrol = cmbRndVF.SelectedItem.ToString();
                        if (sinavKontrol == "Vize")
                        {
                            sinavKontrol = "VİZESİ";
                        }
                        else if (sinavKontrol == "Final")
                        {
                            sinavKontrol = "FİNALİ";
                        }
                        ustbilgi = "MARMARA ÜNİVERSİTESİ TEKNİK BİLİMLER MESLEK YÜKSEKOKULU\n" + txtRndBolum.Text.ToString().ToUpper() + " PROGRAMI " + cmbSinavDers.SelectedItem.ToString().ToUpper() + " DERSİ " + sinavKontrol.ToString() + "\nSINAV SÜRESİ=" + txtRndSure.Text.ToString().ToUpper();
                        headerRangeB.Text = ustbilgi + "\n\nAd:                                            Soyad:                                            No:                                            \nB GURUBU";


                        word.Table oTableB;
                        word.Range wrdRngB = icerikB.Bookmarks.get_Item(ref sonB).Range;
                        oTableB = icerikB.Tables.Add(wrdRngB, tabloSatirSayisi, 1, ref omissingB, ref omissingB);
                        oTableB.Range.ParagraphFormat.SpaceAfter = 6;
                        oTableB.Range.Font.Size = 10;


                        soru = "";
                        a = "";
                        b = "";
                        c = "";
                        d = "";
                        eC = "";
                        //dogru = "";
                        sayac = 0;
                        siklar = "";
                        tumSoru = "";
                        soruResmi = "";

                        for (int i = 1; i <= tabloSatirSayisi; i++)
                        {
                            soru = dtB.Rows[sayac]["soru"].ToString().Trim();
                            a = dtB.Rows[sayac]["a_cevap"].ToString().Trim();
                            b = dtB.Rows[sayac]["b_cevap"].ToString().Trim();
                            c = dtB.Rows[sayac]["c_cevap"].ToString().Trim();
                            d = dtB.Rows[sayac]["d_cevap"].ToString().Trim();
                            eC = dtB.Rows[sayac]["e_cevap"].ToString().Trim();
                            siklar = "a)" + a + "\nb)" + b + "\nc)" + c + "\nd)" + d + "\ne)" + eC + "\n";
                            tumSoru = (sayac + 1).ToString() + "-)" + soru;
                            soruResmi = dtB.Rows[sayac]["soru_resim"].ToString().Trim();
                            if (soruResmi != null && soruResmi != "")
                            {

                                oTableB.Cell(i, 1).Range.InlineShapes.AddPicture((@soruResmi), Type.Missing, Type.Missing, Type.Missing);
                                i++;
                                oTableB.Cell(i, 1).Range.Font.Bold = 1;

                                oTableB.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTableB.Cell(i, 1).Range.Font.Bold = 0;
                                if (dtB.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtB.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        //oTable.Rows[i].Cells[1].SplitCell(3, 2);
                                        oTableB.Rows[i].Cells[1].Split(3, 4);
                                        oTableB.Cell(i, 1).Range.Text = "a)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTableB.Cell(i, 3).Range.Text = "b)";
                                        oTableB.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTableB.Cell(i, 1).Range.Text = "c)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTableB.Cell(i, 3).Range.Text = "d)";
                                        oTableB.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTableB.Cell(i, 1).Range.Text = "e)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTableB.Cell(i, 1).Range.Text = siklar;
                                    }
                                }

                                else
                                {
                                    oTableB.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            else
                            {

                                oTableB.Cell(i, 1).Range.Font.Bold = 1;

                                oTableB.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTableB.Cell(i, 1).Range.Font.Bold = 0;
                                if (dtB.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtB.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        oTableB.Rows[i].Cells[1].Split(3, 4);

                                        oTableB.Cell(i, 1).Range.Text = "a)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTableB.Cell(i, 3).Range.Text = "b)";
                                        oTableB.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTableB.Cell(i, 1).Range.Text = "c)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTableB.Cell(i, 3).Range.Text = "d)";
                                        oTableB.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTableB.Cell(i, 1).Range.Text = "e)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTableB.Cell(i, 1).Range.Text = siklar;
                                    }
                                }
                                else
                                {
                                    oTableB.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            sayac++;

                        }

                        //B Grubu Cevap Anahtarı

                        object omissing3 = System.Reflection.Missing.Value;
                        object son3 = "\\endofdoc";

                        word.Application cevapAnahtariB;
                        word.Document icerik3;
                        cevapAnahtariB = new word.Application();
                        cevapAnahtariB.Visible = true;
                        icerik3 = cevapAnahtariB.Documents.Add(ref omissing3);

                        icerik3.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme

                        /*
                        word.Range headerRangeCevaplarB = icerik3.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplarB.Fields.Add(headerRangeCevaplarB, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplarB.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplarB.Font.Size = 10;
                        headerRangeCevaplarB.Font.Bold = 1;
                        headerRangeCevaplarB.Text = rtxtRndUstbilgi.Text + "\nB GRUBU CEVAP ANAHTARI\n";
                        */

                        word.Range headerRangeCevaplarB = icerik3.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplarB.Fields.Add(headerRangeCevaplarB, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplarB.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplarB.Font.Size = 10;
                        headerRangeCevaplarB.Font.Bold = 1;
                        //string sinavKontrol = cmbRndVF.SelectedItem.ToString();
                        if (sinavKontrol == "Vize")
                        {
                            sinavKontrol = "VİZESİ";
                        }
                        else if (sinavKontrol == "Final")
                        {
                            sinavKontrol = "FİNALİ";
                        }
                        ustbilgi = "MARMARA ÜNİVERSİTESİ TEKNİK BİLİMLER MESLEK YÜKSEKOKULU\n" + txtRndBolum.Text.ToString().ToUpper() + " PROGRAMI " + cmbSinavDers.SelectedItem.ToString().ToUpper() + " DERSİ " + sinavKontrol.ToString() + "\nSINAV SÜRESİ=" + txtRndSure.Text.ToString().ToUpper();
                        headerRangeCevaplarB.Text = ustbilgi + "\nB GURUBU CEVAP ANAHTARI";


                        word.Table oTableBC;
                        word.Range wrdRngBC = icerik3.Bookmarks.get_Item(ref son3).Range;
                        oTableBC = icerik3.Tables.Add(wrdRngBC, tabloSatirSayisi, 1, ref omissing2, ref omissing2);
                        oTableBC.Range.ParagraphFormat.SpaceAfter = 6;
                        oTableBC.Range.Font.Size = 10;

                        soru2 = "";
                        sayac2 = 0;
                        cevapSatirSayisi = soruSayisi * 2;
                        tumSoru2 = "";
                        dogruSik = "";
                        dogruCevap = "";
                        for (int i = 1; i <= cevapSatirSayisi; i++)
                        {
                            soru2 = dtB.Rows[sayac2]["soru"].ToString().Trim();
                            tumSoru2 = (sayac2 + 1).ToString() + "-)" + soru2;
                            dogruCevap = dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim();
                            if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["a_cevap"].ToString().Trim())
                            {
                                dogruSik = "A-)";
                            }
                            else if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["b_cevap"].ToString().Trim())
                            {
                                dogruSik = "B-)";
                            }
                            else if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["c_cevap"].ToString().Trim())
                            {
                                dogruSik = "C-)";
                            }
                            else if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["d_cevap"].ToString().Trim())
                            {
                                dogruSik = "D-)";
                            }
                            else if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["e_cevap"].ToString().Trim())
                            {
                                dogruSik = "E-)";
                            }
                            oTableBC.Cell(i, 1).Range.Font.Bold = 1;
                            oTableBC.Cell(i, 1).Range.Text = tumSoru2;
                            i++;
                            oTableBC.Cell(i, 1).Range.Font.Bold = 0;
                            oTableBC.Cell(i, 1).Range.Text = dogruSik + dogruCevap;
                            sayac2++;

                        }

                    }
                }
                else
                {
                    MessageBox.Show("Lütfen yıldızlı(*) alanları boş bırakmayınız.");
                }

            }
            else
            {
                MessageBox.Show("Seçilen sorular listesinde soru olmalıdır.");
            }
        }

        private void rdbSikResim_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbSikResim.Checked == true)
            {
                txtA.Visible = false;
                txtB.Visible = false;
                txtC.Visible = false;
                txtD.Visible = false;
                txtE.Visible = false;
                btnA.Visible = true;
                btnB.Visible = true;
                btnC.Visible = true;
                btnD.Visible = true;
                btnE.Visible = true;
                pcbA.Visible = true;
                pcbB.Visible = true;
                pcbC.Visible = true;
                pcbD.Visible = true;
                pcbE.Visible = true;
                label54.Visible = true;
                label55.Visible = true;
                label56.Visible = true;
                label57.Visible = true;
                label77.Visible = true;

            }
        }

        private void rdbSikMetin_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbSikMetin.Checked == true)
            {
                txtA.Visible = true;
                txtB.Visible = true;
                txtC.Visible = true;
                txtD.Visible = true;
                txtE.Visible = true;
                btnA.Visible = false;
                btnB.Visible = false;
                btnC.Visible = false;
                btnD.Visible = false;
                btnE.Visible = false;
                pcbA.Visible = false;
                pcbB.Visible = false;
                pcbC.Visible = false;
                pcbD.Visible = false;
                pcbE.Visible = false;
                label54.Visible = false;
                label55.Visible = false;
                label56.Visible = false;
                label57.Visible = false;
                label77.Visible = false;

            }
        }

        private void btnA_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            aResim = dosya.FileName;
            //pcbSoruResmi.ImageLocation = aResim;
        }

        private void btnB_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            bResim = dosya.FileName;
            //pcbSoruResmi.ImageLocation = aResim;
        }

        private void btnC_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            cResim = dosya.FileName;
            //pcbSoruResmi.ImageLocation = aResim;
        }

        private void btnD_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dResim = dosya.FileName;
            //pcbSoruResmi.ImageLocation = aResim;
        }

        private void btnSonuc_Click(object sender, EventArgs e)
        {


            Form5 frm5 = new Form5();
            frm5.Show();

            btnDersEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSoruEklePanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSinavOlusturPanel.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            button1.BackColor = System.Drawing.ColorTranslator.FromHtml("#04345c");
            btnSonuc.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSonuc.ForeColor = System.Drawing.ColorTranslator.FromHtml("#04345c");

            button1.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSinavOlusturPanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnSoruEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");
            btnDersEklePanel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#FFF");

            pnlDersEkle.Visible = false;
            pnlOzelSinav.Visible = false;
            pnlRastgeleOlustur.Visible = false;
            pnlSoruEkle.Visible = false;
        }

        private void btnOzelWord_Click(object sender, EventArgs e)
        {
            if (lstOzelSinav.Items.Count > 0)
            {
                if ((cmbOzelGrup.SelectedItem.ToString() != "" || cmbOzelGrup.SelectedItem != null) && (cmbOzelVF.SelectedItem.ToString() != "" || cmbOzelVF.SelectedItem != null) && txtOzelBolum.Text != "" && txtOzelSure.Text != "" && (cmbOzelDersler.SelectedItem.ToString() != "" || cmbOzelDersler.SelectedItem != null))
                {
                    if (cmbOzelGrup.SelectedItem.ToString() == "Tek Gurup")
                    {
                        int tabloSatirSayisi = 0;
                        int sikSayisi = 0, resimSayisi = 0;
                        int soruSayisi;
                        soruSayisi = dtSecilenSorular.Rows.Count;
                        sikSayisi = soruSayisi;
                        for (int i = 0; i < soruSayisi; i++)
                        {
                            if (dtSecilenSorular.Rows[i]["soru_resim"].ToString() != null && dtSecilenSorular.Rows[i]["soru_resim"].ToString() != "")
                            {
                                resimSayisi++;
                            }
                        }
                        //tabloSatirSayisi = (dtA.Rows.Count * 2) + resimliSorular + (resimliSiklar);
                        tabloSatirSayisi = soruSayisi + sikSayisi + resimSayisi;
                        //tabloSatirSayisi
                        //soruSayisi

                        object omissing = System.Reflection.Missing.Value;
                        object son = "\\endofdoc";

                        word.Application olustur;
                        word.Document icerik;
                        olustur = new word.Application();
                        olustur.Visible = true;
                        icerik = olustur.Documents.Add(ref omissing);

                        icerik.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        //A grubu sayfa üst bilgisi
                        word.Range headerRangeA = icerik.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeA.Fields.Add(headerRangeA, word.WdFieldType.wdFieldPage);
                        headerRangeA.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeA.Font.Size = 10;
                        headerRangeA.Font.Bold = 1;
                        string sinavKontrol = cmbOzelVF.SelectedItem.ToString();
                        if (sinavKontrol == "Vize")
                        {
                            sinavKontrol = "VİZESİ";
                        }
                        else if (sinavKontrol == "Final")
                        {
                            sinavKontrol = "FİNALİ";
                        }
                        string ustbilgi = "MARMARA ÜNİVERSİTESİ TEKNİK BİLİMLER MESLEK YÜKSEKOKULU\n" + txtOzelBolum.Text.ToString().ToUpper() + " PROGRAMI " + cmbOzelDersler.SelectedItem.ToString().ToUpper() + " DERSİ " + sinavKontrol.ToString() + "\nSINAV SÜRESİ=" + txtOzelSure.Text.ToString().ToUpper();
                        headerRangeA.Text = ustbilgi + "\n\nAd:                                            Soyad:                                            No:                                            \nA GRUBU \n";


                        /*word.Range headerRange = icerik.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Fields.Add(headerRange, word.WdFieldType.wdFieldPage);
                        headerRange.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRange.Font.Size = 10;
                        headerRange.Font.Bold = 1;*/
                        //headerRangeB.Text = rtxtRndUstbilgi.Text + "\n\nAd:                                            Soyad:                                            No:                                            \nA GRUBU \n";


                        word.Table oTable;
                        word.Range wrdRng = icerik.Bookmarks.get_Item(ref son).Range;
                        oTable = icerik.Tables.Add(wrdRng, tabloSatirSayisi, 1, ref omissing, ref omissing);
                        oTable.Range.ParagraphFormat.SpaceAfter = 6;
                        oTable.Range.Font.Size = 10;


                        string soru = "";
                        string a = "";
                        string b = "";
                        string c = "";
                        string d = "";
                        string eC = "";
                        //string dogru = "";
                        int sayac = 0;
                        string siklar = "";
                        string tumSoru = "";
                        string soruResmi = "";

                        for (int i = 1; i <= tabloSatirSayisi; i++)
                        {
                            soru = dtSecilenSorular.Rows[sayac]["soru"].ToString().Trim();
                            a = dtSecilenSorular.Rows[sayac]["a_cevap"].ToString().Trim();
                            b = dtSecilenSorular.Rows[sayac]["b_cevap"].ToString().Trim();
                            c = dtSecilenSorular.Rows[sayac]["c_cevap"].ToString().Trim();
                            d = dtSecilenSorular.Rows[sayac]["d_cevap"].ToString().Trim();
                            eC = dtSecilenSorular.Rows[sayac]["e_cevap"].ToString().Trim();
                            siklar = "a)" + a + "\nb)" + b + "\nc)" + c + "\nd)" + d + "\ne)" + eC + "\n";
                            tumSoru = (sayac + 1).ToString() + "-)" + soru;
                            soruResmi = dtSecilenSorular.Rows[sayac]["soru_resim"].ToString().Trim();
                            if (soruResmi != null && soruResmi != "")
                            {

                                oTable.Cell(i, 1).Range.InlineShapes.AddPicture((@soruResmi), Type.Missing, Type.Missing, Type.Missing);
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 1;

                                oTable.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 0;

                                if (dtSecilenSorular.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtSecilenSorular.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        //oTable.Rows[i].Cells[1].SplitCell(3, 2);
                                        oTable.Rows[i].Cells[1].Split(3, 4);
                                        oTable.Cell(i, 1).Range.Text = "a)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "b)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "c)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "d)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;

                                        oTable.Cell(i, 1).Range.Text = "e)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);


                                    }
                                    else
                                    {
                                        oTable.Cell(i, 1).Range.Text = siklar;
                                    }
                                }

                                else
                                {
                                    oTable.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            else
                            {

                                oTable.Cell(i, 1).Range.Font.Bold = 1;

                                oTable.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 0;

                                if (dtSecilenSorular.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtSecilenSorular.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        oTable.Rows[i].Cells[1].Split(2, 4);

                                        oTable.Cell(i, 1).Range.Text = "a)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "b)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "c)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "d)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "e)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTable.Cell(i, 1).Range.Text = siklar;
                                    }
                                }
                                else
                                {
                                    oTable.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            sayac++;

                        }


                        //------A Cevap Anahtarı

                        object omissing2 = System.Reflection.Missing.Value;
                        object son2 = "\\endofdoc";

                        word.Application cevapAnahtari;
                        word.Document icerik2;
                        cevapAnahtari = new word.Application();
                        cevapAnahtari.Visible = true;
                        icerik2 = cevapAnahtari.Documents.Add(ref omissing2);

                        icerik2.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        word.Range headerRangeCevaplar = icerik2.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplar.Fields.Add(headerRangeCevaplar, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplar.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplar.Font.Size = 10;
                        headerRangeCevaplar.Font.Bold = 1;
                        headerRangeCevaplar.Text = ustbilgi + "\nA GRUBU CEVAP ANAHTARI\n";


                        word.Table oTableC;
                        word.Range wrdRngC = icerik2.Bookmarks.get_Item(ref son2).Range;
                        oTableC = icerik2.Tables.Add(wrdRngC, tabloSatirSayisi, 1, ref omissing2, ref omissing2);
                        oTableC.Range.ParagraphFormat.SpaceAfter = 6;
                        oTableC.Range.Font.Size = 10;

                        string soru2 = "";
                        int sayac2 = 0;
                        int cevapSatirSayisi = soruSayisi * 2;
                        string tumSoru2 = "";
                        string dogruSik = "";
                        string dogruCevap = "";
                        for (int i = 1; i <= cevapSatirSayisi; i++)
                        {
                            soru2 = dtSecilenSorular.Rows[sayac2]["soru"].ToString().Trim();
                            tumSoru2 = (sayac2 + 1).ToString() + "-)" + soru2;
                            dogruCevap = dtSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim();
                            if (dtSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtSecilenSorular.Rows[sayac2]["a_cevap"].ToString().Trim())
                            {
                                dogruSik = "A-)";
                            }
                            else if (dtSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtSecilenSorular.Rows[sayac2]["b_cevap"].ToString().Trim())
                            {
                                dogruSik = "B-)";
                            }
                            else if (dtSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtSecilenSorular.Rows[sayac2]["c_cevap"].ToString().Trim())
                            {
                                dogruSik = "C-)";
                            }
                            else if (dtSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtSecilenSorular.Rows[sayac2]["d_cevap"].ToString().Trim())
                            {
                                dogruSik = "D-)";
                            }
                            else if (dtSecilenSorular.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtSecilenSorular.Rows[sayac2]["e_cevap"].ToString().Trim())
                            {
                                dogruSik = "E-)";
                            }
                            oTableC.Cell(i, 1).Range.Font.Bold = 1;
                            oTableC.Cell(i, 1).Range.Text = tumSoru2;
                            i++;
                            oTableC.Cell(i, 1).Range.Font.Bold = 0;
                            oTableC.Cell(i, 1).Range.Text = dogruSik + dogruCevap;
                            sayac2++;





                        }
                    }
                    else if (cmbOzelGrup.SelectedItem.ToString() == "İki Gurup")
                    {


                        DataTable dtA = new DataTable();
                        DataTable dtB = new DataTable();

                        dtA.Columns.Add("soru_id");
                        dtA.Columns.Add("soru");
                        dtA.Columns.Add("a_cevap");
                        dtA.Columns.Add("b_cevap");
                        dtA.Columns.Add("c_cevap");
                        dtA.Columns.Add("d_cevap");
                        dtA.Columns.Add("dogru_cevap");
                        dtA.Columns.Add("zorluk_derecesi");
                        dtA.Columns.Add("soru_resim");

                        dtB.Columns.Add("soru_id");
                        dtB.Columns.Add("soru");
                        dtB.Columns.Add("a_cevap");
                        dtB.Columns.Add("b_cevap");
                        dtB.Columns.Add("c_cevap");
                        dtB.Columns.Add("d_cevap");
                        dtB.Columns.Add("dogru_cevap");
                        dtB.Columns.Add("zorluk_derecesi");
                        dtB.Columns.Add("soru_resim");

                        //dtA.DefaultView.Sort = "soru_id DESC";
                        dtA.Merge(dtSecilenSorular);

                        dtB.Merge(dtSecilenSorular);
                        DataView dw = dtB.DefaultView;
                        dw.Sort = "soru";
                        dtB = dw.ToTable();

                        int tabloSatirSayisi = 0;
                        int sikSayisi = 0, resimSayisi = 0;
                        int soruSayisi;
                        soruSayisi = dtSecilenSorular.Rows.Count;
                        sikSayisi = soruSayisi;
                        for (int i = 0; i < soruSayisi; i++)
                        {
                            if (dtSecilenSorular.Rows[i]["soru_resim"].ToString() != null && dtSecilenSorular.Rows[i]["soru_resim"].ToString() != "")
                            {
                                resimSayisi++;
                            }
                        }
                        //tabloSatirSayisi = (dtA.Rows.Count * 2) + resimliSorular + (resimliSiklar);
                        tabloSatirSayisi = soruSayisi + sikSayisi + resimSayisi;


                        object omissing = System.Reflection.Missing.Value;
                        object son = "\\endofdoc";

                        word.Application olustur;
                        word.Document icerik;
                        olustur = new word.Application();
                        olustur.Visible = true;
                        icerik = olustur.Documents.Add(ref omissing);

                        icerik.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        word.Range headerRangeA = icerik.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeA.Fields.Add(headerRangeA, word.WdFieldType.wdFieldPage);
                        headerRangeA.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeA.Font.Size = 10;
                        headerRangeA.Font.Bold = 1;
                        string sinavKontrol = cmbOzelVF.SelectedItem.ToString();
                        if (sinavKontrol == "Vize")
                        {
                            sinavKontrol = "VİZESİ";
                        }
                        else if (sinavKontrol == "Final")
                        {
                            sinavKontrol = "FİNALİ";
                        }
                        string ustbilgi = "MARMARA ÜNİVERSİTESİ TEKNİK BİLİMLER MESLEK YÜKSEKOKULU\n" + txtOzelBolum.Text.ToString().ToUpper() + " PROGRAMI " + cmbOzelDersler.SelectedItem.ToString().ToUpper() + " DERSİ " + sinavKontrol.ToString() + "\nSINAV SÜRESİ=" + txtOzelSure.Text.ToString().ToUpper() + "\n";
                        headerRangeA.Text = ustbilgi + "\n\nAd:                                            Soyad:                                            No:                                            \nA GRUBU \n";


                        word.Table oTable;
                        word.Range wrdRng = icerik.Bookmarks.get_Item(ref son).Range;
                        oTable = icerik.Tables.Add(wrdRng, tabloSatirSayisi, 1, ref omissing, ref omissing);
                        oTable.Range.ParagraphFormat.SpaceAfter = 6;
                        oTable.Range.Font.Size = 10;


                        string soru = "";
                        string a = "";
                        string b = "";
                        string c = "";
                        string d = "";
                        string eC = "";
                        //string dogru = "";
                        int sayac = 0;
                        string siklar = "";
                        string tumSoru = "";
                        string soruResmi = "";

                        for (int i = 1; i <= tabloSatirSayisi; i++)
                        {
                            soru = dtA.Rows[sayac]["soru"].ToString().Trim();
                            a = dtA.Rows[sayac]["a_cevap"].ToString().Trim();
                            b = dtA.Rows[sayac]["b_cevap"].ToString().Trim();
                            c = dtA.Rows[sayac]["c_cevap"].ToString().Trim();
                            d = dtA.Rows[sayac]["d_cevap"].ToString().Trim();
                            eC = dtA.Rows[sayac]["d_cevap"].ToString().Trim();
                            siklar = "a)" + a + "\nb)" + b + "\nc)" + c + "\nd)" + d + "\nc)" + eC + "\n";
                            tumSoru = (sayac + 1).ToString() + "-)" + soru;
                            soruResmi = dtA.Rows[sayac]["soru_resim"].ToString().Trim();
                            if (soruResmi != null && soruResmi != "")
                            {

                                oTable.Cell(i, 1).Range.InlineShapes.AddPicture((@soruResmi), Type.Missing, Type.Missing, Type.Missing);
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 1;

                                oTable.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 0;

                                if (dtA.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtA.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        //oTable.Rows[i].Cells[1].SplitCell(3, 2);
                                        oTable.Rows[i].Cells[1].Split(3, 4);
                                        oTable.Cell(i, 1).Range.Text = "a)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "b)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "c)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "d)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "e)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTable.Cell(i, 1).Range.Text = siklar;
                                    }
                                }

                                else
                                {
                                    oTable.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            else
                            {

                                oTable.Cell(i, 1).Range.Font.Bold = 1;

                                oTable.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTable.Cell(i, 1).Range.Font.Bold = 0;

                                if (dtA.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtA.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        oTable.Rows[i].Cells[1].Split(2, 4);

                                        oTable.Cell(i, 1).Range.Text = "a)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "b)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "c)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTable.Cell(i, 3).Range.Text = "d)";
                                        oTable.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTable.Cell(i, 1).Range.Text = "e)";
                                        oTable.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);
                                    }
                                    else
                                    {
                                        oTable.Cell(i, 1).Range.Text = siklar;
                                    }
                                }
                                else
                                {
                                    oTable.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            sayac++;

                        }


                        //------A Cevap Anahtarı

                        object omissing2 = System.Reflection.Missing.Value;
                        object son2 = "\\endofdoc";

                        word.Application cevapAnahtari;
                        word.Document icerik2;
                        cevapAnahtari = new word.Application();
                        cevapAnahtari.Visible = true;
                        icerik2 = cevapAnahtari.Documents.Add(ref omissing2);

                        icerik2.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        word.Range headerRangeCevaplar = icerik2.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplar.Fields.Add(headerRangeCevaplar, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplar.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplar.Font.Size = 10;
                        headerRangeCevaplar.Font.Bold = 1;
                        headerRangeCevaplar.Text = ustbilgi + "\nA GRUBU CEVAP ANAHTARI\n";


                        word.Table oTableC;
                        word.Range wrdRngC = icerik2.Bookmarks.get_Item(ref son2).Range;
                        oTableC = icerik2.Tables.Add(wrdRngC, tabloSatirSayisi, 1, ref omissing2, ref omissing2);
                        oTableC.Range.ParagraphFormat.SpaceAfter = 6;
                        oTableC.Range.Font.Size = 10;

                        string soru2 = "";
                        int sayac2 = 0;
                        int cevapSatirSayisi = soruSayisi * 2;
                        string tumSoru2 = "";
                        string dogruSik = "";
                        string dogruCevap = "";
                        for (int i = 1; i <= cevapSatirSayisi; i++)
                        {
                            soru2 = dtA.Rows[sayac2]["soru"].ToString().Trim();
                            tumSoru2 = (sayac2 + 1).ToString() + "-)" + soru2;
                            dogruCevap = dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim();
                            if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["a_cevap"].ToString().Trim())
                            {
                                dogruSik = "A-)";
                            }
                            else if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["b_cevap"].ToString().Trim())
                            {
                                dogruSik = "B-)";
                            }
                            else if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["c_cevap"].ToString().Trim())
                            {
                                dogruSik = "C-)";
                            }
                            else if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["d_cevap"].ToString().Trim())
                            {
                                dogruSik = "D-)";
                            }
                            else if (dtA.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtA.Rows[sayac2]["e_cevap"].ToString().Trim())
                            {
                                dogruSik = "E-)";
                            }
                            oTableC.Cell(i, 1).Range.Font.Bold = 1;
                            oTableC.Cell(i, 1).Range.Text = tumSoru2;
                            i++;
                            oTableC.Cell(i, 1).Range.Font.Bold = 0;
                            oTableC.Cell(i, 1).Range.Text = dogruSik + dogruCevap;
                            sayac2++;





                        }
                        // B GRUBU

                        object omissingB = System.Reflection.Missing.Value;
                        object sonB = "\\endofdoc";

                        word.Application olusturB;
                        word.Document icerikB;
                        olusturB = new word.Application();
                        olusturB.Visible = true;
                        icerikB = olusturB.Documents.Add(ref omissingB);

                        icerikB.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        word.Range headerRangeB = icerikB.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeB.Fields.Add(headerRangeB, word.WdFieldType.wdFieldPage);
                        headerRangeB.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeB.Font.Size = 10;
                        headerRangeB.Font.Bold = 1;
                        headerRangeB.Text = ustbilgi + "\n\nAd:                                            Soyad:                                            No:                                            \nB GRUBU \n";


                        word.Table oTableB;
                        word.Range wrdRngB = icerikB.Bookmarks.get_Item(ref sonB).Range;
                        oTableB = icerikB.Tables.Add(wrdRngB, tabloSatirSayisi, 1, ref omissingB, ref omissingB);
                        oTableB.Range.ParagraphFormat.SpaceAfter = 6;
                        oTableB.Range.Font.Size = 10;


                        soru = "";
                        a = "";
                        b = "";
                        c = "";
                        d = "";
                        eC = "";
                        //dogru = "";
                        sayac = 0;
                        siklar = "";
                        tumSoru = "";
                        soruResmi = "";

                        for (int i = 1; i <= tabloSatirSayisi; i++)
                        {
                            soru = dtB.Rows[sayac]["soru"].ToString().Trim();
                            a = dtB.Rows[sayac]["a_cevap"].ToString().Trim();
                            b = dtB.Rows[sayac]["b_cevap"].ToString().Trim();
                            c = dtB.Rows[sayac]["c_cevap"].ToString().Trim();
                            d = dtB.Rows[sayac]["d_cevap"].ToString().Trim();
                            eC = dtB.Rows[sayac]["e_cevap"].ToString().Trim();
                            siklar = "a)" + a + "\nb)" + b + "\nc)" + c + "\nd)" + d + "\ne)" + eC + "\n";
                            tumSoru = (sayac + 1).ToString() + "-)" + soru;
                            soruResmi = dtB.Rows[sayac]["soru_resim"].ToString().Trim();
                            if (soruResmi != null && soruResmi != "")
                            {

                                oTableB.Cell(i, 1).Range.InlineShapes.AddPicture((@soruResmi), Type.Missing, Type.Missing, Type.Missing);
                                i++;
                                oTableB.Cell(i, 1).Range.Font.Bold = 1;

                                oTableB.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTableB.Cell(i, 1).Range.Font.Bold = 0;
                                if (dtB.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtB.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        //oTable.Rows[i].Cells[1].SplitCell(3, 2);
                                        oTableB.Rows[i].Cells[1].Split(3, 4);
                                        oTableB.Cell(i, 1).Range.Text = "a)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTableB.Cell(i, 3).Range.Text = "b)";
                                        oTableB.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTableB.Cell(i, 1).Range.Text = "c)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTableB.Cell(i, 3).Range.Text = "d)";
                                        oTableB.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTableB.Cell(i, 1).Range.Text = "e)";
                                        //MessageBox.Show(eC.ToString());
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTableB.Cell(i, 1).Range.Text = siklar;
                                    }
                                }

                                else
                                {
                                    oTableB.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            else
                            {

                                oTableB.Cell(i, 1).Range.Font.Bold = 1;

                                oTableB.Cell(i, 1).Range.Text = tumSoru;
                                i++;
                                oTableB.Cell(i, 1).Range.Font.Bold = 0;
                                if (dtB.Rows[sayac]["a_cevap"].ToString().Length >= 3)
                                {
                                    if (dtB.Rows[sayac]["a_cevap"].ToString().Trim().Substring(0, 3) == @"C:\")
                                    {
                                        oTableB.Rows[i].Cells[1].Split(2, 4);

                                        oTableB.Cell(i, 1).Range.Text = "a)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@a), Type.Missing, Type.Missing, Type.Missing);

                                        oTableB.Cell(i, 3).Range.Text = "b)";
                                        oTableB.Cell(i, 4).Range.InlineShapes.AddPicture((@b), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTableB.Cell(i, 1).Range.Text = "c)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@c), Type.Missing, Type.Missing, Type.Missing);

                                        oTableB.Cell(i, 3).Range.Text = "d)";
                                        oTableB.Cell(i, 4).Range.InlineShapes.AddPicture((@d), Type.Missing, Type.Missing, Type.Missing);
                                        i++;
                                        oTableB.Cell(i, 1).Range.Text = "e)";
                                        oTableB.Cell(i, 2).Range.InlineShapes.AddPicture((@eC), Type.Missing, Type.Missing, Type.Missing);

                                    }
                                    else
                                    {
                                        oTableB.Cell(i, 1).Range.Text = siklar;
                                    }
                                }
                                else
                                {
                                    oTableB.Cell(i, 1).Range.Text = siklar;
                                }
                            }
                            sayac++;

                        }

                        //B Grubu Cevap Anahtarı

                        object omissing3 = System.Reflection.Missing.Value;
                        object son3 = "\\endofdoc";

                        word.Application cevapAnahtariB;
                        word.Document icerik3;
                        cevapAnahtariB = new word.Application();
                        cevapAnahtariB.Visible = true;
                        icerik3 = cevapAnahtariB.Documents.Add(ref omissing3);

                        icerik3.PageSetup.TextColumns.SetCount(2);// Sayfayı ikiye bölme


                        word.Range headerRangeCevaplarB = icerik3.Sections[1].Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRangeCevaplarB.Fields.Add(headerRangeCevaplarB, word.WdFieldType.wdFieldPage);
                        headerRangeCevaplarB.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphCenter;

                        headerRangeCevaplarB.Font.Size = 10;
                        headerRangeCevaplarB.Font.Bold = 1;
                        headerRangeCevaplarB.Text = ustbilgi + "\nB GRUBU CEVAP ANAHTARI\n";


                        word.Table oTableBC;
                        word.Range wrdRngBC = icerik3.Bookmarks.get_Item(ref son3).Range;
                        oTableBC = icerik3.Tables.Add(wrdRngBC, tabloSatirSayisi, 1, ref omissing2, ref omissing2);
                        oTableBC.Range.ParagraphFormat.SpaceAfter = 6;
                        oTableBC.Range.Font.Size = 10;

                        soru2 = "";
                        sayac2 = 0;
                        cevapSatirSayisi = soruSayisi * 2;
                        tumSoru2 = "";
                        dogruSik = "";
                        dogruCevap = "";
                        for (int i = 1; i <= cevapSatirSayisi; i++)
                        {
                            soru2 = dtB.Rows[sayac2]["soru"].ToString().Trim();
                            tumSoru2 = (sayac2 + 1).ToString() + "-)" + soru2;
                            dogruCevap = dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim();
                            if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["a_cevap"].ToString().Trim())
                            {
                                dogruSik = "A-)";
                            }
                            else if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["b_cevap"].ToString().Trim())
                            {
                                dogruSik = "B-)";
                            }
                            else if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["c_cevap"].ToString().Trim())
                            {
                                dogruSik = "C-)";
                            }
                            else if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["d_cevap"].ToString().Trim())
                            {
                                dogruSik = "D-)";
                            }
                            else if (dtB.Rows[sayac2]["dogru_cevap"].ToString().Trim() == dtB.Rows[sayac2]["e_cevap"].ToString().Trim())
                            {
                                dogruSik = "E-)";
                            }
                            oTableBC.Cell(i, 1).Range.Font.Bold = 1;
                            oTableBC.Cell(i, 1).Range.Text = tumSoru2;
                            i++;
                            oTableBC.Cell(i, 1).Range.Font.Bold = 0;
                            oTableBC.Cell(i, 1).Range.Text = dogruSik + dogruCevap;
                            sayac2++;





                        }
                    }
                }
                else
                {
                    MessageBox.Show("Lütfen yıldızlı(*) alanları boş bırakmayınız.");
                }


            }
            else
            {
                MessageBox.Show("Lütfen soru seçimi yapınız.");
            }
        }

        private void btnE_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            eResim = dosya.FileName;
        }

        private void txtSinavKolay_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar >= 47 && (int)e.KeyChar <= 58)
            {

                e.Handled = false;

            }

            else if ((int)e.KeyChar == 8)
            {

                e.Handled = false;

            }

            else
            {

                e.Handled = true;

            }

        }

        private void txtSinavOrta_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar >= 47 && (int)e.KeyChar <= 58)
            {

                e.Handled = false;

            }

            else if ((int)e.KeyChar == 8)
            {

                e.Handled = false;

            }

            else
            {

                e.Handled = true;

            }

        }

        private void txtSinavZor_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar >= 47 && (int)e.KeyChar <= 58)
            {

                e.Handled = false;

            }

            else if ((int)e.KeyChar == 8)
            {

                e.Handled = false;

            }

            else
            {

                e.Handled = true;

            }

        }
    }
}
