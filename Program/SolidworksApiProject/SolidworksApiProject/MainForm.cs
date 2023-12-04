using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace SolidworksApiProject
{
        
    public partial class MainForm : Form
    {
        public static SldWorks swApp;

        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter5.Chapter5Form c5f = new Chapter5.Chapter5Form();
            c5f.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter6.Chapter6Form c6f = new Chapter6.Chapter6Form();
            c6f.ShowDialog();
            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter7.Chapter7Form c7f = new Chapter7.Chapter7Form();
            c7f.ShowDialog();
            this.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter8.Chapter8Form c8f = new Chapter8.Chapter8Form();
            c8f.ShowDialog();
            this.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter9.Chapter9Form c9f = new Chapter9.Chapter9Form();
            c9f.ShowDialog();
            this.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter10.Chapter10Form c10f = new Chapter10.Chapter10Form();
            c10f.ShowDialog();
            this.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter11.Chapter11Form c11f = new Chapter11.Chapter11Form();
            c11f.ShowDialog();
            this.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter12.Chapter12Form c12f = new Chapter12.Chapter12Form();
            c12f.ShowDialog();
            this.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter13.Chapter13Form c13f = new Chapter13.Chapter13Form();
            c13f.ShowDialog();
            this.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter14.Chapter14Form c14f = new Chapter14.Chapter14Form();
            c14f.ShowDialog();
            this.Show();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            this.Hide();
            Chapter15.Chapter15Form c15f = new Chapter15.Chapter15Form();
            c15f.ShowDialog();
            this.Show();
        }

     


    }
}
