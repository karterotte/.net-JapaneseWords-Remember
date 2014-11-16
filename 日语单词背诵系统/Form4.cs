using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Data.OleDb;
using System.Collections;//ArrayList才可以使用

using System.Drawing.Design;  //控件的附加功能
using System.Drawing.Drawing2D; //图形容器、混合、高级画笔、矩阵、变形
using System.Drawing.Imaging;  //基本图像处理功能
using System.Drawing.Text;  //

namespace 日语单词背诵系统
{

    public partial class Form4 : Form
    {
        ArrayList selectedDates = new ArrayList();//想要在各窗体之间传递数据，定义static全局变量即可
        DataTable conditions = new DataTable();

        OleDbDataAdapter myAda = new OleDbDataAdapter();
        DataSet myDs = new DataSet();

        public Form4()
        {
            InitializeComponent();
            selectedDates = Form3.selectedDates;
          
        }

        public  void StartConn2()
        {
            string sqlstr = "select 单词组,开始日期,第一次日期,第一次完成情况,第二次日期,第二次完成情况,第三次日期,第三次完成情况,第四次日期,第四次完成情况,完成情况 from 计划总表 ";

            try
            {
                Conn conn = new Conn();
                OleDbConnection connection = conn.CreatConn();
                connection.Open();
                myAda = new OleDbDataAdapter(sqlstr, connection);


                //myAda.Fill(Words);

                myAda.Fill(myDs, "myWords");
                conditions = myDs.Tables["myWords"];

                //DataTable myWords = new DataTable();
                //myAda.Fill(myWords);//不用myDs，直接导入到表中即可
                connection.Close();

            }
            catch
            {
                MessageBox.Show("检查数据库联接！");
                return;
            }
        }

        //public void Listadd()
        //{
        //    listView1.Items.Clear();

        //    //BeginUpdate()的意思是停止listview的更新,以免加载内容时不停的闪烁
        //    listView1.BeginUpdate();
        //    for (int i = 0; i < Words.Rows.Count; i++)
        //    {
        //        //item加载的是每行的第一列内容
        //        //对应的item的属性subitem加载的是后面列的内容


        //        ListViewItem Myitems = new ListViewItem(Words.Rows[i]["平假名"].ToString());
        //        Myitems.SubItems.Add(Words.Rows[i]["汉字"].ToString());
        //        Myitems.SubItems.Add(Words.Rows[i]["释义"].ToString());
        //        listView1.Items.Add(Myitems);

        //    }
        //    listView1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent);
        //    //AutoResizeColumn放在数据加载前面没有效果,如果是根据内容确定的话.
        //    listView1.EndUpdate();
        //    //EndUpdate()的意思是listview开始更新了
        //}



        private void Form4_Load(object sender, EventArgs e)
        {
            StartConn2();
            if (conditions.Rows.Count == 0)
            {

            }
            else
            {
                this.tabPage1.Text  = conditions.Rows[0]["开始日期"].ToString();
                System.Windows.Forms.Button bb1 = new Button(); ;
                bb1.Location = new System.Drawing.Point(513, 40);
                bb1.Size = new System.Drawing.Size(75, 23);
                bb1.Text = "FUXK";
                bb1 .UseVisualStyleBackColor = true;
                bb1.Visible = true;
                this.tabPage1.Controls.Add(bb1);

            }
        }

 










    }
}
