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
{
    public partial class Dialog2 : Form
    {

        string dancibiao = form1.G_danci;
        string jihuabiao = form1.G_jihua;

        public Dialog2()
        {
            InitializeComponent();
            
        }

        public void readQuery(string sqlstr, OleDbConnection oleConnection, CheckedListBox checkedListBox,bool key)
        {
            checkedListBox.Items.Clear();
            try
            {
                OleDbCommand myCommand = new OleDbCommand(sqlstr, oleConnection);
                OleDbDataReader myReader = myCommand.ExecuteReader();
                while (myReader.Read())
                {
                    if (!checkedListBox.Items.Contains(myReader.GetString(0)))
                    {
                        checkedListBox.Items.Add(myReader.GetString(0));
                    }
                }
                myReader.Close();
                oleConnection.Close();

                if (checkedListBox.Items.Count == 0 && key ==true )
                {
                    MessageBox.Show("现在没有单词记录，请录入先！");
                    return;
                }
            }
            catch
            {
                MessageBox.Show("发生错误", "提示！", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);

            }
        }




        private void button2_Click(object sender, EventArgs e)
        {
            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();
            string sqlstring = "select distinct 描述,录入日期 from " + dancibiao + " order by 录入日期 desc;";//录入日期
            readQuery(sqlstring, connection, checkedListBox1,true );
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form3.selectedDates.Clear();
            string[] rqs={};

            if (checkedListBox1.CheckedItems .Count  == 0)
            {
                MessageBox.Show("现在没有单词，请录入或选择先！");
                return ;
            }

            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                Console.WriteLine(checkedListBox1.CheckedItems[i].ToString());

                Form3.selectedDates.Add(checkedListBox1.CheckedItems[i].ToString());
                //调用全局变量，必须加前面的类

            }
            Close();
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string dayvalue = DateTime.Now.Date.ToString("yyyy-MM-dd");
            //DateTime dayvalue = Convert.ToDateTime(dayvalue2);
            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();
            string sqlstring = "select 描述 from " + jihuabiao + " where 第一次日期 = '" + dayvalue + "' or  第二次日期 = '" + dayvalue + "' or  第三次日期 = '" + dayvalue + "' or  第四次日期 = '" + dayvalue + "' ;";//单词组



            readQuery(sqlstring, connection, checkedListBox1,false );
            connection.Close();
            if (checkedListBox1.Items.Count == 0)
            {
                MessageBox.Show("今天没有计划任务!");
            }
        }


    }
}
