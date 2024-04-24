using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Properties.Methods;
using Properties.Objects;


namespace Properties
{
    public partial class Form1 : Form
    {
        public int maxBar;
        public Form1()
        {
            InitializeComponent();
            this.maxBar = ReadExcel.validItens;
            progressBar1.Maximum = maxBar;
        }

        public int count = 0;
        public void UpdateBarProgress()
        {
            int step = 1;

            if (progressBar1.Value < progressBar1.Maximum)
                progressBar1.Value += step;

        }

        public void killForms()
        {
            this.Close();
        }
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }
    }
}
