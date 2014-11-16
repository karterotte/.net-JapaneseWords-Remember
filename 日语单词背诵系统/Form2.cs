using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections;


namespace 日语单词背诵系统
{     //form2和form1同在一个namespace中，定义的东西可以不用using就通用

    
    public partial class Form2 : Form
    {

        //事先要声明才能被类中其他函数捕捉
        private DataSet InsertWords;
        private DataTable Words;
        public  string desc = "";
        public static bool key = true;
        
        public Form2()
        {
            InitializeComponent();

            //全局的东西就放在类的初始化中，事先要声明才能被类中其他函数捕捉

            InsertWords = new DataSet("单词录入"); //进入本窗体，实例化一个dataset，等保存退出时候再释放
            Words = InsertWords.Tables.Add("单词表");
            Words.Columns.Add("假名", typeof(string));
            Words.Columns.Add("汉字", typeof(string));
            Words.Columns.Add("释义", typeof(string));
            Words.Columns.Add("词性", typeof(string));
            Words.Columns.Add("等级", typeof(string));
            Words.Columns.Add("例句", typeof(string));
            Words.Columns.Add("日期", typeof(string));
            Words.Columns.Add("描述", typeof(string));
            comboBox1.Text = form1.wt;
            comboBox2.Text = form1.wl;
            strIn si = new strIn();
            si.ShowDialog();
            desc = form1.desc;

        }

        public void Clear()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            //comboBox1.Text = "";
            //comboBox2.Text = "";
        }
        public void AddRecords()
        {
            DataRow Wordrows = Words.NewRow();
            //每次需要改变的变量放在其他函数中
            string anon="";
            if (textBox5.Text == "")
            {

            }
            else
            {
                anon = "  【" + textBox5.Text.ToString().Trim() + "】";
            }

            Wordrows["假名"] = textBox1.Text.ToString()+anon ;
            Wordrows["汉字"] = textBox2.Text.ToString();
            Wordrows["释义"] = textBox3.Text.ToString();
            Wordrows["例句"] = textBox4.Text.ToString();
            Wordrows["词性"] = comboBox1.Text.ToString();
            Wordrows["等级"] = comboBox2.Text.ToString();
            Wordrows["日期"] = DateTime.Now.ToString("yyyy年MM月dd日");//获取当前日期而不得到时间
                                                                       //以文本格式规范日期  
            Wordrows["描述"] = desc;

            Words.Rows.Add(Wordrows);

            this.Clear();//类中的函数和变量 可以直接调用，也可以加上this强调

            int count = InsertWords.Tables["单词表"].Rows.Count + 1;
            groupBox1.Text = count.ToString() + "/50";//获取表中记录总数

            comboBox1.Text = form1 .wt ;
            comboBox2.Text = form1.wl;

            DataRow kk = Wordrows;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox3.Text == "")
            {
                //-------------------------------------------------------------------
                //若有错误发生errorProvider事件！
                if (textBox1.Text == "")
                {
                    errorProvider1.SetError(textBox1, "不能为空！");
                }
                if (textBox3.Text == "")
                {
                    errorProvider1.SetError(textBox3, "不能为空！");
                }
                return;
                //-------------------------------------------------------------------
            }

            this.AddRecords();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clear();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                this.AddRecords();//最后上传时如果还有一条记录没有加载！则再执行一次“保存”
            }



            Conn conn=new Conn ();
            OleDbConnection connection=conn.CreatConn ();
            connection .Open ();
            int count = InsertWords.Tables["单词表"].Rows.Count;
            string dancibiao = form1.G_danci;
            string sqlstring = "insert into "+dancibiao+"(平假名,汉字,释义,词性,单词级别,录入日期,例句,描述)values";
            string valuestring = "";
            for (int i = 0; i < count; i++)
            {
                string jiaming = InsertWords.Tables["单词表"].Rows[i]["假名"].ToString();
                string hanzi = InsertWords.Tables["单词表"].Rows[i]["汉字"].ToString();
                string shiyi = InsertWords.Tables["单词表"].Rows[i]["释义"].ToString();
                string cixing = InsertWords.Tables["单词表"].Rows[i]["词性"].ToString();
                string dengji = InsertWords.Tables["单词表"].Rows[i]["等级"].ToString();
                string riqi = InsertWords.Tables["单词表"].Rows[i]["日期"].ToString();
                string liju = InsertWords.Tables["单词表"].Rows[i]["例句"].ToString();
                string miaoshu = InsertWords.Tables["单词表"].Rows[i]["描述"].ToString();
                valuestring = "('" + jiaming + "','" + hanzi + "','" + shiyi + "','" + cixing + "','" + dengji + "','" + riqi + "','" + liju + "','"+miaoshu +"');";
                
                //insert语句中，values值为文本的列的值都必须加单引号括起来！！！！

                
                OleDbCommand comm = new OleDbCommand(sqlstring + valuestring, connection);
                comm.ExecuteNonQuery();

            }


            connection.Close();
            MessageBox.Show("上传成功!");
            Close();

        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            trackBar1.Maximum = Words.Rows.Count;
            trackBar1.Minimum = 0;
            int i = trackBar1.Value;

        }

        private void Form2_Activated(object sender, EventArgs e)
        {
            if (key == false )
            {
                this.Close();
            }
        }







        
    }
}
