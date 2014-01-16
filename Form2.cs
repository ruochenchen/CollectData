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
            this.textBox_Destination.Text = fileWay;
        }

        private void button_Scan_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            openFileDialog.Filter = "Excel文件|*.xls*";
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox_Source.Text = openFileDialog.FileName;
            }
        }

        private void button_Comfirm_Click(object sender, EventArgs e)
        {
            if (this.textBox_Source.Text != "")
            {
                fileWay_Source = this.textBox_Source.Text;
                this.Close();
            }
            else
            {
                MessageBox.Show("请点击浏览按钮选择源数据文件所在的路径！");
            }
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
