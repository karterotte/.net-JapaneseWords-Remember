using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace 日语单词背诵系统
{
    public partial class Form9 : Form
    {
        private int restTime = 10;

        public Form9()
        {
            InitializeComponent();
        }

        private void Form9_Load(object sender, EventArgs e)
        {
            timer1.Interval = 1000;
            timer1.Start();
            button1.Text = restTime.ToString();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (restTime == 0)
            {
                timer1.Stop();
                button1.Text = "Ready Go!";
            }
            else
            {
                restTime -= 1;
                button1.Text = restTime.ToString ();
            }
        }

        
    }
}
