using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace checkdataCollect
{
    public partial class Form2 : Form
    {
        public string fileWay_Source;

        public Form2()
        {
            InitializeComponent();
        }

        public Form2(string fileWay)
        {
            InitializeComponent();
            textBox1.Text = fileWay;
        }

        private void button_Scan_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            openFileDialog.Filter = "Excel文件|*.xls*";
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = openFileDialog.FileName;
            }
        }

        private void button_Comfirm_Click(object sender, EventArgs e)
        {
            if (this.textBox2.Text != "")
            {
                fileWay_Source = this.textBox2.Text;
            }
            this.Close();
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
