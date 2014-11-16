using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data .OleDb ;
using DotNetSpeech;
using System.IO;

namespace 日语单词背诵系统
{

    
    
    public partial class form1 : Form
    {
        public static string[] WordType;
        public static string wt;                //录入默认词汇类型

        public static string[] WordLevel;
        public static string wl;                //录入默认词汇级别

        public static bool Yinbiao = true ;

        public static string G_danci = "";      //单词表目标
        public static string G_jihua = "";      //计划表目标

        public static string voice = "";        //语音角色目标

        public static string desc = "";         //录入单词描述

        public form1()
        {
            InitializeComponent();
            try
            {
                FileInfo file = new FileInfo("OOXX.ini");
                StreamReader reader = file.OpenText();
                string text = reader.ReadLine();
                wt=  text.Split(',')[0];
                wl = text.Split(',')[1];

                string text2 = reader.ReadLine();
                G_danci = text2.Split(',')[0];
                G_jihua = text2.Split(',')[1];

                voice  = reader.ReadLine();
                reader.Close();
                

            }
            catch
            {
                MessageBox.Show("OOXX文件失效！");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();
            form5.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form6 form6 = new Form6();
            form6.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //SpVoiceClass voice = new SpVoiceClass();
            //voice.Voice = voice.GetVoices("name=ScanSoft Kyoko_Full_22kHz", string.Empty).Item(0);
            ////voice.AllowAudioOutputFormatChangesOnNextSet = true;
            //voice.Rate = -5;
            //voice.Speak(textBox1.Text, SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak);

            //Form4 form4 = new Form4();
            //form4.Show();
            Form10 form10 = new Form10();
            form10.Show();

        }



        private void button6_Click_1(object sender, EventArgs e)
        {
            Form7 form7 = new Form7();
            form7.Show();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            Form8 form8 = new Form8();
            form8.Show();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            Form9 form9 = new Form9();
            form9.Show();
        }

        private void form1_Load(object sender, EventArgs e)
        {
            int iresult;
            Random ro = new Random();
            iresult = ro.Next(1, 11); 
            switch (iresult)
            {
                case 1: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._1; break;
                case 2: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._2; break;
                case 3: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._3; break;
                case 4: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._4; break;
                case 5: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._5; break;
                case 6: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._6; break;
                case 7: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._7; break;
                case 8: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._8; break;
                case 9: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._9; break;
                case 10: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._10; break;
                default: this.pictureBox1.Image = global::日语单词背诵系统.Properties.Resources._1; break;
            }
            
        }


    }

    class Conn
    {
        private string datasouce;
        public Conn()
        {
            //datasouce = "F:\\数据库\\JapaneseWords.mdb";
            datasouce = Application.StartupPath + "\\JapaneseWords.mdb";

            //加载进来之后就可以直接使用名字，不需要带路径了
            //datasouce = " JapaneseWords.mdb";
        }
        public OleDbConnection CreatConn()
        {
            OleDbConnection conn = new OleDbConnection
            ("provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + datasouce + ";");
            Console.WriteLine("provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + datasouce + ";");
            return conn;
        }

    }

   
}
