using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace Sinav_Olusturma
{
    public partial class Form5 : Form
    {
        double soruPuani;
        //int soruSayisi = 0;
        int[] soruAnalizA;
        int[] soruAnalizB;
        //int sB = 0;
        int[] aralikA = new int[20];
        System.Data.DataTable dtSoruAnalizA = new System.Data.DataTable();
        System.Data.DataTable dtSoruAnalizB = new System.Data.DataTable();
        System.Data.DataTable dtAralik = new System.Data.DataTable();
        int grupSayisi = 0;
        public Form5()
        {
            InitializeComponent();
        }
        public void notAraligi(double x)
        {
            if (x >= 0 && x < 5)
            {
                aralikA[0] += 1;
            }
            else if (x >= 5 && x < 10)
            {
                aralikA[1] += 1;
            }
            else if (x >= 10 && x < 15)
            {
                aralikA[2] += 1;
            }
            else if (x >= 15 && x < 20)
            {
                aralikA[3] += 1;
            }
            else if (x >= 20 && x < 25)
            {
                aralikA[4] += 1;
            }
            else if (x >= 25 && x < 30)
            {
                aralikA[5] += 1;
            }
            else if (x >= 30 && x < 35)
            {
                aralikA[6] += 1;
            }
            else if (x >= 35 && x < 40)
            {
                aralikA[7] += 1;
            }
            else if (x >= 40 && x < 45)
            {
                aralikA[8] += 1;
            }
            else if (x >= 45 && x < 50)
            {
                aralikA[9] += 1;
            }
            else if (x >= 50 && x < 55)
            {
                aralikA[10] += 1;
            }
            else if (x >= 55 && x < 60)
            {
                aralikA[11] += 1;
            }
            else if (x >= 60 && x < 65)
            {
                aralikA[12] += 1;
            }
            else if (x >= 65 && x < 70)
            {
                aralikA[13] += 1;
            }
            else if (x >= 70 && x < 75)
            {
                aralikA[14] += 1;
            }
            else if (x >= 75 && x < 80)
            {
                aralikA[15] += 1;
            }
            else if (x >= 80 && x < 85)
            {
                aralikA[16] += 1;
            }
            else if (x >= 85 && x < 90)
            {
                aralikA[17] += 1;
            }
            else if (x >= 90 && x < 95)
            {
                aralikA[18] += 1;
            }
            else if (x >= 95 && x <= 100)
            {
                aralikA[19] += 1;
            }
        }
        private void Form5_Load(object sender, EventArgs e)
        {
            dtSoruAnalizA.Columns.Add("Soru Numarası");
            dtSoruAnalizA.Columns.Add("Doğru Cevaplayan Sayısı");
            dtSoruAnalizB.Columns.Add("Soru Numarası");
            dtSoruAnalizB.Columns.Add("Doğru Cevaplayan Sayısı");
            dtAralik.Columns.Add("Aralık");
            dtAralik.Columns.Add("Kişi Sayısı");
        }

        private void Form5_DragDrop(object sender, DragEventArgs e)
        {
            grupSayisi = 0;
            lblACevaplar.Text = "...";
            lblBCevaplar.Text = "...";
            lblSoruPuani.Text = "...";
            dtSoruAnalizA.Rows.Clear();
            dtSoruAnalizB.Rows.Clear();
            dtAralik.Rows.Clear();
            dataGridView1.Rows.Clear();
            string[] veriler = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            StreamReader aktar = new StreamReader(veriler[0], Encoding.GetEncoding("iso-8859-9"), false);
            string yazi;
            //int grupBul;
            int soruSayisi = 0;
            string aCevaplar = "", bCevaplar = "";
            //richTextBox1.Text = aktar.ReadToEnd();
            //int sayac = 0;
            string ogrNo, ogrAd, sinavTur, grup, ogCevaplar;
            int dogru = 0, yanlis = 0, bos = 0;
            double puan = 0;
            while ((yazi = aktar.ReadLine()) != null)
            {
                if (yazi.Substring(0, 5).ToString() == "00000")
                {

                    if (yazi.Substring(10, 1).ToString() == "A")
                    {
                        //MessageBox.Show("A Grubu");
                        aCevaplar = yazi.Substring(35, yazi.Length - 35).Trim();
                        soruSayisi = aCevaplar.Length;
                        soruPuani = 100.00 / soruSayisi;
                        lblSoruPuani.Text = soruPuani.ToString();
                        lblACevaplar.Text = aCevaplar;
                        grupSayisi++;
                        //MessageBox.Show("A gurubu cevapları bulundu");
                    }
                }
                if (yazi.Substring(0, 5).ToString() == "00000")
                {

                    if (yazi.Substring(10, 1).ToString() == "B")
                    {
                        //MessageBox.Show("B Grubu");
                        bCevaplar = yazi.Substring(35, yazi.Length - 35).Trim();
                        lblBCevaplar.Text = bCevaplar;
                        //MessageBox.Show("B gurubu cevapları bulundu");
                        grupSayisi++;
                    }
                }



            }
            MessageBox.Show("Sınavdaki gurup sayısı=" + grupSayisi + ". Cevap anahtarı tespit edildi");

            /**/
            if (grupSayisi == 1)
            {
                soruAnalizA = new int[soruSayisi];
                for (int i = 0; i < soruSayisi; i++)
                {
                    soruAnalizA[i] = 0;
                }
            }
            if (grupSayisi == 2)
            {
                soruAnalizA = new int[soruSayisi];
                soruAnalizB = new int[soruSayisi];
                for (int i = 0; i < soruSayisi; i++)
                {
                    soruAnalizA[i] = 0;
                    soruAnalizB[i] = 0;
                }
            }
            /**/
            if (grupSayisi <= 1)
            {
                lblABaslik.Text = "A gurubundaki cevaplar=";
                lblBBaslik.Text = "...";
                lblBCevaplar.Text = "...";
            }
            else
            {
                lblABaslik.Text = "A gurubundaki cevaplar=";
                lblBBaslik.Text = "B gurubundaki cevaplar=";
            }
            string[] veriler2 = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            StreamReader aktar2 = new StreamReader(veriler[0], Encoding.GetEncoding("iso-8859-9"), false);

            string yazi2;
            while ((yazi2 = aktar2.ReadLine()) != null)
            {
                //MessageBox.Show("While a girdi");
                if (yazi2.Substring(0, 5).ToString() != "00000")
                {
                    //MessageBox.Show("ife girdi");
                    ogrNo = yazi2.Substring(0, 8).ToString();
                    sinavTur = yazi2.Substring(9, 1).ToString();
                    if (sinavTur == "1")
                    {
                        sinavTur = "Vize";
                    }
                    else
                    {
                        sinavTur = "Final";
                    }
                    grup = yazi2.Substring(10, 1).ToString();
                    ogrAd = yazi2.Substring(11, 25).ToString().Trim();
                    ogCevaplar = yazi2.Substring(36, soruSayisi);
                    //MessageBox.Show("Öğrencinin cevap sayısı=" + ogCevaplar.Length.ToString());
                    /*MessageBox.Show(ogCevaplar.Length.ToString());
                    MessageBox.Show(aCevaplar.Length.ToString());
                    MessageBox.Show(bCevaplar.Length.ToString());*/
                    if (grup == "A")
                    {
                        /*MessageBox.Show(ogrAd);
                        MessageBox.Show(ogCevaplar.Length.ToString());*/
                        for (int i = 0; i < aCevaplar.Length; i++)
                        {

                            if (ogCevaplar.Substring(i, 1) != " ")
                            {
                                if (ogCevaplar.Substring(i, 1) == aCevaplar.Substring(i, 1))
                                {
                                    dogru++;
                                    soruAnalizA[i] += 1;
                                }
                                else
                                {
                                    yanlis++;
                                }
                            }
                            else
                            {
                                bos++;
                            }
                        }
                        puan = Convert.ToDouble(dogru) * soruPuani;
                        notAraligi(puan);
                        dataGridView1.Rows.Add(ogrNo, ogrAd, sinavTur, grup, dogru.ToString(), yanlis.ToString(), bos.ToString(), puan.ToString());
                        //richTextBox1.Text = richTextBox1.Text + ogCevaplar;
                        dogru = 0;
                        yanlis = 0;
                        bos = 0;
                        ogrAd = "";
                        ogCevaplar = "";
                        ogrNo = "";
                        sinavTur = "";
                        grup = "";
                        puan = 0;
                        //MessageBox.Show("Öğrenci Eklendi");
                    }
                    else
                    {
                        for (int i = 0; i < bCevaplar.Length; i++)
                        {
                            //MessageBox.Show(i.ToString());
                            if (ogCevaplar.Substring(i, 1) != " ")
                            {
                                if (ogCevaplar.Substring(i, 1) == bCevaplar.Substring(i, 1))
                                {
                                    //MessageBox.Show(ogCevaplar.Substring(i, 1));

                                    dogru++;
                                    soruAnalizB[i] += 1;
                                }
                                else
                                {
                                    yanlis++;
                                }
                            }
                            else
                            {
                                bos++;
                            }
                        }
                        puan = Convert.ToDouble(dogru) * soruPuani;
                        notAraligi(puan);
                        dataGridView1.Rows.Add(ogrNo, ogrAd, sinavTur, grup, dogru.ToString(), yanlis.ToString(), bos.ToString(), puan.ToString());
                        //richTextBox1.Text = richTextBox1.Text + ogCevaplar;
                        dogru = 0;
                        yanlis = 0;
                        bos = 0;
                        ogrAd = "";
                        ogCevaplar = "";
                        ogrNo = "";
                        sinavTur = "";
                        grup = "";
                        puan = 0;
                        //MessageBox.Show("Öğrenci Eklendi");
                    }
                }
            }
            //MessageBox.Show("A 1. soruyu doğru cevaplayan sayısı= " + soruAnalizA[0].ToString());
            //MessageBox.Show("B 1. soruyu doğru cevaplayan sayısı= " + soruAnalizB[1].ToString());

            int notlar = 0;
            for (int i = 0; i < 20; i++)
            {
                if (notlar == 95)
                {
                    dtAralik.Rows.Add((notlar.ToString() + "-" + (notlar + 5).ToString()), aralikA[i]);
                    notlar += 5;
                }
                else
                {
                    dtAralik.Rows.Add((notlar.ToString() + "-" + (notlar + 4).ToString()), aralikA[i]);
                    notlar += 5;
                }

            }
            if (grupSayisi == 1)
            {
                for (int i = 0; i < soruSayisi; i++)
                {
                    dtSoruAnalizA.Rows.Add((i + 1), soruAnalizA[i]);
                }
            }
            else
            {
                for (int i = 0; i < soruSayisi; i++)
                {
                    dtSoruAnalizA.Rows.Add((i + 1), soruAnalizA[i]);
                    dtSoruAnalizB.Rows.Add((i + 1), soruAnalizB[i]);
                }
            }
            /*for (int i = 0; i < dtSoruAnalizB.Rows.Count; i++)
            {
                MessageBox.Show(dtSoruAnalizB.Rows[i][1].ToString());
            }*/
            //MessageBox.Show(grupSayisi.ToString());
        }

        private void Form5_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void btnExcelAktar_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(grupSayisi.ToString());
            if (dataGridView1.Rows.Count >= 2)
            {
                Excel.Application excelDosya = new Excel.Application();
                

                object Missing = Type.Missing;
                Workbook calismakitabi = excelDosya.Workbooks.Add(Missing);
                Worksheet sheet1 = (Worksheet)calismakitabi.Sheets[1];
                sheet1.Name = "Bireysel Notlar";


                Worksheet sheet2 = (Worksheet)calismakitabi.Sheets.Add(Missing, Missing, 1, Missing) as Excel.Worksheet; ;
                sheet2.Name = "Soru Analizi";

                Worksheet sheet3 = (Worksheet)calismakitabi.Sheets.Add(Missing, Missing, 1, Missing) as Excel.Worksheet; ;
                sheet3.Name = "Not Aralıkları";
                Excel.Range formatRange;

                formatRange = sheet1.get_Range("h:h");
                formatRange.NumberFormat = "@";

                Excel.Range formatRange2;

                formatRange2 = sheet3.get_Range("a:a");
                formatRange2.NumberFormat = "@";
                //sheet1.Cells[8, 2].NumberFormat = "General";
                int sutun = 1;
                int satir = 1;

                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Range myrange = (Range)sheet1.Cells[satir, sutun + j];
                    myrange.Value2 = dataGridView1.Columns[j].HeaderText;
                }
                satir++;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        Range myrange = (Range)sheet1.Cells[satir + i, sutun + j];
                        myrange.Value2 = dataGridView1[j, i].ValueType == null ? "" : dataGridView1[j, i].Value;
                        //myrange.Select();
                    }
                }
                //soru analiz

                if (grupSayisi == 1)
                {
                    int analizSatir = 1, analizSutun = 1;

                    for (int i = 0; i < dtSoruAnalizA.Rows.Count; i++)
                    {
                        for (int j = 0; j < dtSoruAnalizA.Columns.Count; j++)
                        {
                            Range myrange2 = (Range)sheet2.Cells[analizSatir + i, analizSutun + j];
                            myrange2.Value2 = dtSoruAnalizA.Rows[i][j] == null ? "" : dtSoruAnalizA.Rows[i][j];
                            //myrange2.Select();

                        }
                    }
                }
                if (grupSayisi == 2)
                {
                    //MessageBox.Show("Else çalışıyor");
                    int analizSatirA = 1, analizSutunA = 1, analizSatirB = 1, analizSutunB = 4;

                    for (int i = 0; i < dtSoruAnalizA.Rows.Count; i++)
                    {
                        for (int j = 0; j < dtSoruAnalizA.Columns.Count; j++)
                        {
                            Range myrange2 = (Range)sheet2.Cells[analizSatirA + i, analizSutunA + j];
                            myrange2.Value2 = dtSoruAnalizA.Rows[i][j] == null ? "" : dtSoruAnalizA.Rows[i][j];
                            //myrange2.Select();

                        }
                    }

                    for (int i = 0; i < dtSoruAnalizB.Rows.Count; i++)
                    {
                        for (int j = 0; j < dtSoruAnalizB.Columns.Count; j++)
                        {
                            Range myrange2 = (Range)sheet2.Cells[analizSatirB + i, analizSutunB + j];
                            myrange2.Value2 = dtSoruAnalizB.Rows[i][j] == null ? "" : dtSoruAnalizB.Rows[i][j];
                            //myrange2.Select();

                        }
                    }
                }


                //Not aralıkları
                int aralikSatir = 1, araliksutun = 1;
                for (int i = 0; i < dtAralik.Rows.Count; i++)
                {
                    for (int j = 0; j < dtAralik.Columns.Count; j++)
                    {
                        Range myrange3 = (Range)sheet3.Cells[aralikSatir + i, araliksutun + j];
                        myrange3.Value2 = dtAralik.Rows[i][j] == null ? "" : dtAralik.Rows[i][j];
                        //myrange2.Select();

                    }
                }


                //Grafikler



                if (grupSayisi == 1)
                {
                    Excel.Range chartRange;

                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet2.ChartObjects(Type.Missing);
                    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(50, 20, 800, 250);
                    Excel.Chart chartPage = myChart.Chart;

                    chartRange = sheet2.get_Range("B:A");
                    chartPage.SetSourceData(chartRange, Missing);
                    chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

                    excelDosya.Visible = true;
                }
                else
                {
                    Excel.Range chartRange;

                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)sheet2.ChartObjects(Type.Missing);
                    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(50, 20, 800, 250);
                    Excel.Chart chartPage = myChart.Chart;

                    chartRange = sheet2.get_Range("B:A");
                    chartPage.SetSourceData(chartRange, Missing);
                    chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

                    Excel.Range chartRange2;

                    Excel.ChartObjects xlCharts2 = (Excel.ChartObjects)sheet2.ChartObjects(Type.Missing);
                    Excel.ChartObject myChart2 = (Excel.ChartObject)xlCharts2.Add(50, 240, 800, 250);
                    Excel.Chart chartPage2 = myChart2.Chart;

                    chartRange2 = sheet2.get_Range("D:E");
                    chartPage2.SetSourceData(chartRange2, Missing);
                    chartPage2.ChartType = Excel.XlChartType.xlColumnClustered;

                    excelDosya.Visible = true;
                }


                //Not Aralıkları

                Excel.Range chartRange3;

                Excel.ChartObjects xlCharts3 = (Excel.ChartObjects)sheet3.ChartObjects(Type.Missing);
                Excel.ChartObject myChart3 = (Excel.ChartObject)xlCharts3.Add(50, 80, 800, 250);
                Excel.Chart chartPage3 = myChart3.Chart;

                chartRange3 = sheet3.get_Range("A:B");
                chartPage3.SetSourceData(chartRange3, Missing);
                chartPage3.ChartType = Excel.XlChartType.xlColumnClustered;

                excelDosya.Visible = true;
            }
            else
            {
                MessageBox.Show("Aktarmaya uygun veri bulunamadı.");
            }


        }
    }
}
