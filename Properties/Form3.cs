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

namespace Properties
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            progressBar1.Maximum = ReadExcel.rowCounts;
        }

        public void UpdateProgressBar()
        {
            int step = 1;

            if (progressBar1.Value < progressBar1.Maximum)
                progressBar1.Value += step;

            
        }


        public void killForms()
        {
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
