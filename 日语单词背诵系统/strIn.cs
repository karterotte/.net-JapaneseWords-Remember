using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace 日语单词背诵系统
{
    public partial class strIn : Form
    {
        public strIn()
        {
            InitializeComponent();
        }

        private void strIn_Load(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.ToString("yyyy年MM月dd日");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            form1.desc = textBox1.Text;
            Form2.key = true;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2.key = false;
            this.Close();
        }
    }
}
