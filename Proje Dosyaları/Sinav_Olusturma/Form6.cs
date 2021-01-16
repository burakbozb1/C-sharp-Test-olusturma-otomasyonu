using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Sinav_Olusturma
{
    public partial class Form6 : Form
    {
        int sayac = 0;
        public Form6()
        {
            InitializeComponent();
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            //pictureBox1.ImageLocation = "logo.png";
            this.Opacity = 1.0;
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            sayac++;
            if (sayac > 100)
            {
                if (this.Opacity == 0.00)
                {
                    MessageBox.Show("test");
                    timer1.Stop();
                    Form1 frm1 = new Form1();
                    frm1.Show();
                    this.Hide();
                }
                else
                {
                    this.Opacity -= 0.01;
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            
        }

        private void Form6_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        private void Form6_GiveFeedback(object sender, GiveFeedbackEventArgs e)
        {

        }

        private void Form6_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }
    }
}
