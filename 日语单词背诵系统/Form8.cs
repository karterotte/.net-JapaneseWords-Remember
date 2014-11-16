using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections;//ArrayList才可以使用
using System.IO;
using DotNetSpeech;

namespace 日语单词背诵系统
{
    public partial class Form8 : Form
    {
        string dancibiao = form1.G_danci;
        string jihuabiao = form1.G_jihua;
        string voiced = form1.voice;

        public Form8()
        {
            InitializeComponent();
            comboBox1.Text = form1.wt;
            comboBox2.Text = form1.wl;

            SpVoiceClass voice = new SpVoiceClass();
            int ss = voice.GetVoices(string.Empty, string.Empty).Count;
            string kk = "";
            comboBox_voice.Items.Clear();
            for (int nn = 0; nn < ss; nn++)
            {
                kk = (string)voice.GetVoices(string.Empty, string.Empty).Item(nn).GetDescription(0);
                comboBox_voice.Items.Add(kk);
            }


            if (form1.G_danci == "单词总表")
            {
                radioButton3.Checked = true;
            }
            else
            {
                radioButton4.Checked = true;
            }
            comboBox_voice.Text = voiced;
        }



        private void buttonX1_Click(object sender, EventArgs e)
        {
            string sqlstr = "";
            string sqlDate = "";
            if (radioButton2.Checked == true)
            {
                //加载多个日期段的组
                if (Form3.selectedDates.Count == 0)
                {
                    MessageBox.Show("没有选择单词组！");
                    radioButton1.Checked = true;
                    return;
                }
                else if (Form3.selectedDates.Count == 1)
                {
                    sqlDate = ("'" + Form3.selectedDates[0].ToString() + "')");
                }
                else
                {
                    sqlDate = ("'" + Form3.selectedDates[0].ToString() + "'");
                    for (int i = 1; i < Form3.selectedDates.Count; i++)
                    {
                        sqlDate += ("or  录入日期 ='" + Form3.selectedDates[i].ToString() + "'");
                    }
                    sqlDate = sqlDate + ")";
                }
            }
            if (sqlDate != "")
            {
                sqlstr += "UPDATE "+dancibiao +" SET 背诵标志 = false where (录入日期 =" + sqlDate+";";
            }
            else
            {
                sqlstr = "UPDATE " + dancibiao + " SET 背诵标志 = false;";
            }

            try
            {
                Conn conn = new Conn();
                OleDbConnection connection = conn.CreatConn();
                connection.Open();
                OleDbCommand cmd = new OleDbCommand(sqlstr, connection);
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("OK！");
            }
            catch
            {
                MessageBox.Show("检查数据库联接！");
                return;
            }

        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            string sqlstr = "";
            string sqlDate = "";
            if (radioButton2.Checked == true)
            {
                //加载多个日期段的组
                if (Form3.selectedDates.Count == 0)
                {
                    MessageBox.Show("没有选择单词组！");
                    radioButton1.Checked = true;
                    return;
                }
                else if (Form3.selectedDates.Count == 1)
                {
                    sqlDate = ("'" + Form3.selectedDates[0].ToString() + "')");
                }
                else
                {
                    sqlDate = ("'" + Form3.selectedDates[0].ToString() + "'");
                    for (int i = 1; i < Form3.selectedDates.Count; i++)
                    {
                        sqlDate += ("or  录入日期 ='" + Form3.selectedDates[i].ToString() + "'");
                    }
                    sqlDate = sqlDate + ")";
                }
            }
            if (sqlDate != "")
            {
                sqlstr += "UPDATE " + dancibiao + " SET 背诵重要级别 = 1 where (录入日期 =" + sqlDate + ";";
            }
            else
            {
                sqlstr = "UPDATE " + dancibiao + " SET 背诵重要级别 = 1;";
            }


            try
            {
                Conn conn = new Conn();
                OleDbConnection connection = conn.CreatConn();
                connection.Open();
                OleDbCommand cmd = new OleDbCommand(sqlstr, connection);
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("OK！");
            }
            catch
            {
                MessageBox.Show("检查数据库联接！");
                return;
            }
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            string sqlstr = "";
            string sqlDate = "";
            if (radioButton2.Checked == true)
            {
                //加载多个日期段的组
                if (Form3.selectedDates.Count == 0)
                {
                    MessageBox.Show("没有选择单词组！");
                    radioButton1.Checked = true;
                    return;
                }
                else if (Form3.selectedDates.Count == 1)
                {
                    sqlDate = ("'" + Form3.selectedDates[0].ToString() + "')");
                }
                else
                {
                    sqlDate = ("'" + Form3.selectedDates[0].ToString() + "'");
                    for (int i = 1; i < Form3.selectedDates.Count; i++)
                    {
                        sqlDate += ("or  录入日期 ='" + Form3.selectedDates[i].ToString() + "'");
                    }
                    sqlDate = sqlDate + ")";
                }
            }
            if (sqlDate != "")
            {
                sqlstr += "UPDATE " + dancibiao + " SET 总背诵次数 = 0 where (录入日期 =" + sqlDate + ";";
            }
            else
            {
                sqlstr = "UPDATE " + dancibiao + " SET 总背诵次数 = 0;";
            }


            try
            {
                Conn conn = new Conn();
                OleDbConnection connection = conn.CreatConn();
                connection.Open();
                OleDbCommand cmd = new OleDbCommand(sqlstr, connection);
                cmd.ExecuteNonQuery();
                connection.Close();
                MessageBox.Show("OK！");
            }
            catch
            {
                MessageBox.Show("检查数据库联接！");
                return;
            }
        }

        private void buttonX4_Click(object sender, EventArgs e)
        {
            try
            {
                FileInfo file = new FileInfo("OOXX.ini");
                StreamReader reader = file.OpenText();
                string text = reader.ReadLine();
                string wt = text.Split(',')[0];
                string wl = text.Split(',')[1];

                string text2 = reader.ReadLine();
                dancibiao = text2.Split(',')[0];
                jihuabiao = text2.Split(',')[1];

                string voice = reader.ReadLine();
                reader.Close();


                StreamWriter writer = new StreamWriter("OOXX.ini");


                wt = comboBox1.Text;
                wl = comboBox2.Text;
                writer.WriteLine(wt + ',' + wl);

                writer.WriteLine(dancibiao + ',' + jihuabiao);
                writer.WriteLine(voice);
                writer.Close();

                form1.wt = wt;
                form1.wl = wl;

                MessageBox.Show("OK！");

            }
            catch
            {
                MessageBox.Show("OOXX文件失效！");
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                Form3.selectedDates.Clear();
                Dialog2 dil2 = new Dialog2();
                dil2.Show();

            }
            else
            {

            }
        }

        private void buttonX5_Click(object sender, EventArgs e)
        {
            try
            {
                FileInfo file = new FileInfo("OOXX.ini");
                StreamReader reader = file.OpenText();
                string text = reader.ReadLine();
                string wt = text.Split(',')[0];
                string wl = text.Split(',')[1];

                string text2 = reader.ReadLine();
                dancibiao = text2.Split(',')[0];
                jihuabiao = text2.Split(',')[1];

                string voice = reader.ReadLine();
                reader.Close();


                StreamWriter writer = new StreamWriter("OOXX.ini");

                writer.WriteLine(wt + ',' + wl);

                if (radioButton3.Checked)
                {
                    dancibiao = "单词总表";
                    jihuabiao = "计划总表";
                }
                if (radioButton4.Checked)
                {
                    dancibiao = "单词总表2";
                    jihuabiao = "计划总表2";
                }

                writer.WriteLine(dancibiao + ',' + jihuabiao);
                writer.WriteLine(voice);
                writer.Close();

                form1.G_danci = dancibiao;
                form1.G_jihua = jihuabiao;

                MessageBox.Show("OK！");

            }
            catch
            {
                MessageBox.Show("OOXX文件失效！");
            }
        }

        private void buttonX6_Click(object sender, EventArgs e)
        {
            string sel_voice = comboBox_voice.SelectedItem .ToString ();
            

            try
            {
                FileInfo file = new FileInfo("OOXX.ini");
                StreamReader reader = file.OpenText();
                string text = reader.ReadLine();
                string wt = text.Split(',')[0];
                string wl = text.Split(',')[1];

                string text2 = reader.ReadLine();
                dancibiao = text2.Split(',')[0];
                jihuabiao = text2.Split(',')[1];

                string voice = reader.ReadLine();

                reader.Close();


                StreamWriter writer = new StreamWriter("OOXX.ini");

                writer.WriteLine(wt + ',' + wl);

                writer.WriteLine(dancibiao + ',' + jihuabiao);

                if (sel_voice == "")
                {
                    sel_voice = voice;
                }

                writer.WriteLine(sel_voice);

                writer.Close();

                form1.voice = sel_voice;

                MessageBox.Show("OK！");

            }
            catch
            {
                MessageBox.Show("OOXX文件失效！");
            }
        }

        
    }
}
