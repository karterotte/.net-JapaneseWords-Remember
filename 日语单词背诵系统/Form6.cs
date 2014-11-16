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
using System.Windows;

namespace 日语单词背诵系统
{
    public partial class Form6 : Form
    {
        ArrayList selectedDates = new ArrayList();//想要在各窗体之间传递数据，定义static全局变量即可
        DataTable Words = new DataTable();
        DataTable Plans = new DataTable();

        OleDbDataAdapter myAda = new OleDbDataAdapter();
        OleDbDataAdapter myAda2 = new OleDbDataAdapter();
        DataSet myDs = new DataSet();
        DataSet myDs2 = new DataSet();

        string dancibiao = form1.G_danci;
        string jihuabiao = form1.G_jihua;


        public Form6()
        {
            InitializeComponent();
            selectedDates = Form3.selectedDates;
            listView1.Columns.Add("     单词组", 220, HorizontalAlignment.Center);
            listView1.Columns.Add("单词数", 70, HorizontalAlignment.Center);
            listView1.Columns.Add("计划情况", 120, HorizontalAlignment.Center);
            listView1.FullRowSelect = true;

            //listView2.Columns.Add("     单词组", 120, HorizontalAlignment.Center);
            //listView2.Columns.Add("　㈠　", 70, HorizontalAlignment.Center);
            //listView2.Columns.Add("　㈡　", 70, HorizontalAlignment.Center);
            //listView2.Columns.Add("　㈢　", 70, HorizontalAlignment.Center);
            //listView2.Columns.Add("　㈣　", 70, HorizontalAlignment.Center);
            listView2.FullRowSelect = true;
        }

        public void WordsListadd()
        {
            //string sqlString = "select  录入日期,count(录入日期) as 记录数 ,所属计划组 from "+dancibiao +" group by 录入日期 ,所属计划组;";
            string sqlString = "select  描述,count(描述) as 记录数 ,所属计划组 from " + dancibiao + " group by 描述 ,所属计划组;";

            try
            {
                Conn conn = new Conn();
                OleDbConnection connection = conn.CreatConn();
                connection.Open();
                myAda = new OleDbDataAdapter(sqlString, connection);



                myAda.Fill(myDs, "myWords");
                Words = myDs.Tables["myWords"];

                //DataTable myWords = new DataTable();
                //myAda.Fill(myWords);//不用myDs，直接导入到表中即可
                connection.Close();

            }
            catch
            {
                MessageBox.Show("检查数据库联接！");
                return;
            }

            Refresh_Words();
        }
        public void PlanListadd()
        {
            string sqlString = "select  * from "+jihuabiao +";";

            try
            {
                Conn conn = new Conn();
                OleDbConnection connection = conn.CreatConn();
                connection.Open();
                myAda2 = new OleDbDataAdapter(sqlString, connection);



                myAda2.Fill(myDs2, "myPlans");

                Plans = myDs2.Tables["myPlans"];

                //DataTable myWords = new DataTable();
                //myAda.Fill(myWords);//不用myDs，直接导入到表中即可
                connection.Close();

            }
            catch
            {
                MessageBox.Show("检查数据库联接！");
                return;
            }


            Refresh_Plans();
        }

        private void Refresh_Words()
        {
            listView1.Items.Clear();

            //BeginUpdate()的意思是停止listview的更新,以免加载内容时不停的闪烁
            listView1.BeginUpdate();
            for (int i = 0; i < Words.Rows.Count; i++)
            {
                //item加载的是每行的第一列内容
                //对应的item的属性subitem加载的是后面列的内容


                ListViewItem Myitems = new ListViewItem(Words.Rows[i]["描述"].ToString());//录入日期
                Myitems.SubItems.Add(Words.Rows[i]["记录数"].ToString());
                Myitems.SubItems.Add(Words.Rows[i]["所属计划组"].ToString());
                listView1.Items.Add(Myitems);

            }
            listView1.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent);

            //listView1.AutoResizeColumn(2, ColumnHeaderAutoResizeStyle.ColumnContent);
            //AutoResizeColumn放在数据加载前面没有效果,如果是根据内容确定的话.
            listView1.EndUpdate();

            //EndUpdate()的意思是listview开始更新了
        }

        private void Refresh_Plans()
        {
            listView2.Items.Clear();

            //BeginUpdate()的意思是停止listview的更新,以免加载内容时不停的闪烁
            listView2.BeginUpdate();
            for (int i = 0; i < Plans.Rows.Count; i++)
            {
                //item加载的是每行的第一列内容
                //对应的item的属性subitem加载的是后面列的内容


                ListViewItem Myitems = new ListViewItem(Plans.Rows[i]["描述"].ToString());//单词组
                Font ft=new Font ("msyh.ttf",10.5f);
                if (Plans.Rows[i]["第一次完成情况"].ToString() == "False")
                {
                    Myitems.SubItems.Add(Plans.Rows[i]["第一次完成情况"].ToString(), Color.Red, Color.White, ft);
                }
                else
                {
                    Myitems.SubItems.Add(Plans.Rows[i]["第一次完成情况"].ToString(), Color.Green , Color.White, ft);
                }

                
                Myitems.SubItems.Add(Plans.Rows[i]["第二次完成情况"].ToString());
                Myitems.SubItems.Add(Plans.Rows[i]["第三次完成情况"].ToString());
                Myitems.SubItems.Add(Plans.Rows[i]["第四次完成情况"].ToString());
                Console.WriteLine(Plans.Rows[i]["第一次完成情况"].ToString());
                if (Myitems.SubItems[2].Text == "False")
                {
                    Console.WriteLine("111111111111");
                    Myitems.SubItems[2].BackColor = Color.Salmon;
                    //Myitems.ForeColor = Color.DarkBlue;
                    Myitems.SubItems[2].ForeColor = Color.Red;

                }
                listView2.Items.Add(Myitems);

            }
            //listView2.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent);
            //AutoResizeColumn放在数据加载前面没有效果,如果是根据内容确定的话.
            listView2.EndUpdate();
            //EndUpdate()的意思是listview开始更新了
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            WordsListadd();
            PlanListadd();
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            //listView1.SelectedItems.ToString();
            MessageBox.Show(listView1.SelectedItems[0].Text);


        }

        private void button3_Click(object sender, EventArgs e)
        {
            string goal = "";
            try
            {
                goal = listView1.SelectedItems[0].SubItems[2].Text;
            }

            catch
            {
                MessageBox.Show("选定一个项目先！");
                return;
            }
            if (listView1.SelectedItems[0].SubItems[2].Text == "0" | listView1.SelectedItems[0].SubItems[2].Text == "")
            {
                string messageBoxText = "确定要把单词组 【" + listView1.SelectedItems[0].Text + "】  加入到计划中吗？";
                string caption = "法肯达";
                MessageBoxButtons button = MessageBoxButtons.OKCancel;
                MessageBoxIcon icon = MessageBoxIcon.Warning;

                if ("OK" == MessageBox.Show(messageBoxText, caption, button, icon).ToString())
                {
                    string dayvalue = string.Format("{0:D}", DateTime.Now.Date);
                    for (int i = 0; i <= Words.Rows.Count; i++)
                    {
                        if (Words.Rows[i]["描述"].ToString() == listView1.SelectedItems[0].Text)//录入日期
                        {
                            Words.Rows[i]["所属计划组"] = dayvalue;
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    DataRow newrow = Plans.NewRow();
                    newrow["描述"] = listView1.SelectedItems[0].Text;//录入日期
                    newrow["开始日期"] = dayvalue;

                    newrow["第一次"] = 1;
                    newrow["第一次日期"] = DateTime.Now.Date.AddDays(1).ToString("yyyy-MM-dd");
                    newrow["第一次完成情况"] = false;

                    newrow["第二次"] = 3;
                    newrow["第二次日期"] = DateTime.Now.Date.AddDays(3).ToString("yyyy-MM-dd");
                    newrow["第二次完成情况"] = false;

                    newrow["第三次"] = 7;
                    newrow["第三次日期"] = DateTime.Now.Date.AddDays(7).ToString("yyyy-MM-dd");
                    newrow["第三次完成情况"] = false;

                    newrow["第四次"] = 14;
                    newrow["第四次日期"] = DateTime.Now.Date.AddDays(14).ToString("yyyy-MM-dd");
                    newrow["第四次完成情况"] = false;

                    newrow["完成情况"] = false;
                    Plans.Rows.Add(newrow);

                    Refresh_Plans();
                    Refresh_Words();

                }
                else
                {
                    return;
                }

            }
            else
            {
                MessageBox.Show("已加入计划！");
            }
        }

        private void listView1_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            e.DrawBackground();
            //判断Subitem中是否存在关键字
            if (e.SubItem.Text == "0")  //txtContent.Text.Trim().Length > 0 &&
            {
                e.SubItem.ForeColor = Color.Red;  //设置背景色为粉红色
            }
            else
            {
                e.SubItem.ForeColor = Color.Black; //设置字体为红色
            }
            e.DrawText();
        }

        private void listView1_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.DrawBackground();
            e.DrawText();

        }

        private void listView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            e.DrawBackground();
            e.DrawText();
        }

        private void listView2_ItemActivate(object sender, EventArgs e)
        {
            
            for (int i = 0; i < Plans.Rows.Count; i++)
            {
                if (Plans.Rows[i]["描述"].ToString() == listView2.SelectedItems[0].Text)//单词组
                {
                    StartDate.Text = string.Format("{0:D}", Plans.Rows[i]["开始日期"]);
                    numericUpDown1.Value = Convert.ToDecimal(Plans.Rows[i]["第一次"].ToString());
                    numericUpDown2.Value = Convert.ToDecimal(Plans.Rows[i]["第二次"].ToString());
                    numericUpDown3.Value = Convert.ToDecimal(Plans.Rows[i]["第三次"].ToString());
                    numericUpDown4.Value = Convert.ToDecimal(Plans.Rows[i]["第四次"].ToString());

                    DateTime dt = Convert.ToDateTime(StartDate.Text);
                    dateTimePicker1.Value = dt.AddDays(Convert.ToDouble(numericUpDown1.Value));
                    dateTimePicker2.Value = dt.AddDays(Convert.ToDouble(numericUpDown2.Value));
                    dateTimePicker3.Value = dt.AddDays(Convert.ToDouble(numericUpDown3.Value));
                    dateTimePicker4.Value = dt.AddDays(Convert.ToDouble(numericUpDown4.Value));

                    checkBox1.Checked = !Convert.ToBoolean(Plans.Rows[i]["第一次完成情况"]);
                    checkBox2.Checked = !Convert.ToBoolean(Plans.Rows[i]["第二次完成情况"]);
                    checkBox3.Checked = !Convert.ToBoolean(Plans.Rows[i]["第三次完成情况"]);
                    checkBox4.Checked = !Convert.ToBoolean(Plans.Rows[i]["第四次完成情况"]);
                    groupBox3.Refresh();
                    return;
                }
                
            }

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(StartDate.Text);
            dateTimePicker1.Value = dt.AddDays(Convert.ToDouble(numericUpDown1.Value));
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(StartDate.Text);
            TimeSpan ts = dateTimePicker1.Value - dt;
            numericUpDown1.Value = Convert.ToDecimal(ts.Days);

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(StartDate.Text);
            dateTimePicker2.Value = dt.AddDays(Convert.ToDouble(numericUpDown2.Value));
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(StartDate.Text);
            TimeSpan ts = dateTimePicker2.Value - dt;
            numericUpDown2.Value = Convert.ToDecimal(ts.Days);
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(StartDate.Text);
            dateTimePicker3.Value = dt.AddDays(Convert.ToDouble(numericUpDown3.Value));
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(StartDate.Text);
            TimeSpan ts = dateTimePicker3.Value - dt;
            numericUpDown3.Value = Convert.ToDecimal(ts.Days);
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(StartDate.Text);
            dateTimePicker4.Value = dt.AddDays(Convert.ToDouble(numericUpDown4.Value));
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = Convert.ToDateTime(StartDate.Text);
            TimeSpan ts = dateTimePicker4.Value - dt;
            numericUpDown4.Value = Convert.ToDecimal(ts.Days);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string goal = "";
            try
            {
                goal = listView2.SelectedItems[0].Text;
            }
            catch
            {
                MessageBox.Show("选择一个项目先！");
            }
            for (int i = 0; i < Plans.Rows.Count; i++)
            {
                if (Plans.Rows[i]["描述"].ToString() == goal)//单词组
                {
                    Plans.Rows[i]["开始日期"] = StartDate.Text;
                    Plans.Rows[i]["第一次"] = numericUpDown1.Value;
                    Plans.Rows[i]["第二次"] = numericUpDown2.Value;
                    Plans.Rows[i]["第三次"] = numericUpDown3.Value;
                    Plans.Rows[i]["第四次"] = numericUpDown4.Value;

                    Plans.Rows[i]["第一次日期"] = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    Plans.Rows[i]["第二次日期"] = dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    Plans.Rows[i]["第三次日期"] = dateTimePicker3.Value.ToString("yyyy-MM-dd");
                    Plans.Rows[i]["第四次日期"] = dateTimePicker4.Value.ToString("yyyy-MM-dd");

                    if (!checkBox1.Checked)
                    {
                        Plans.Rows[i]["第一次完成情况"] = true;
                    }
                    else
                    {
                        Plans.Rows[i]["第一次完成情况"] = false;
                    }
                    if (!checkBox2.Checked)
                    {
                        Plans.Rows[i]["第二次完成情况"] = true;
                    }
                    else
                    {
                        Plans.Rows[i]["第二次完成情况"] = false;

                    }
                    if (!checkBox3.Checked)
                    {
                        Plans.Rows[i]["第三次完成情况"] = true;
                    }
                    else
                    {
                        Plans.Rows[i]["第三次完成情况"] = false;
                    }

                    if (!checkBox4.Checked)
                    {
                        Plans.Rows[i]["第四次完成情况"] = true;
                    }
                    else
                    {
                        Plans.Rows[i]["第四次完成情况"] = false;
                    }

                    break;
                }
            }


            //OleDbCommandBuilder db = new OleDbCommandBuilder(myAda);
            //myAda.Update(Words);

            OleDbCommandBuilder db2 = new OleDbCommandBuilder(myAda2);
           
            myAda2.Update(Plans);
            Refresh_Plans();

            try
            {
                Conn conn = new Conn();
                OleDbConnection connection = conn.CreatConn();
                connection.Open();
                string updatestring = "update " + dancibiao + " set 所属计划组= '" + StartDate.Text + "' where 描述= '" + goal + "';";//录入日期
                Console.WriteLine(updatestring);
                OleDbCommand ucom = new OleDbCommand(updatestring, connection);
                ucom.ExecuteNonQuery();
                connection.Close();
            }
            catch
            {
                MessageBox.Show("连接数据库失败！");
                return;
            }


            Refresh_Words();
            MessageBox.Show("保存好了！");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string query = "";
            try
            {
                query = "确定要删除该计划:  " + listView2.SelectedItems[0].Text + "  ?";
            }
            catch
            {
                MessageBox.Show("选定一个项目先！");
                return;
            }
            string goal = listView2.SelectedItems[0].Text;

            if ("Yes" == MessageBox.Show(query ,"??",MessageBoxButtons.YesNo).ToString ())
            {
                for (int i = 0; i <= Words.Rows.Count; i++)
                {
                    if (Words.Rows[i]["描述"].ToString() == goal)//录入日期
                    {
                        Words.Rows[i]["所属计划组"] = 0;
                        //listView1.Items[i].SubItems["所属计划组"].Text  = "0";
                        break;
                    }
                }
                Refresh_Words();

                for (int i = 0; i < Plans.Rows.Count; i++)
                {
                    if (Plans.Rows[i]["描述"].ToString() == goal)//单词组
                    {
                        Plans.Rows[i].Delete();
                        listView2.Items[i].Remove();
                        break;
                    }

                }
                //Refresh_Plans();

                



                OleDbCommandBuilder db2 = new OleDbCommandBuilder(myAda2);
                myAda2.Update(Plans);


                try
                {
                    Conn conn = new Conn();
                    OleDbConnection connection = conn.CreatConn();
                    connection.Open();
                    string updatestring = "update " + dancibiao + " set 所属计划组= 0 where 描述= '" + goal + "';";//录入日期
                    OleDbCommand ucom = new OleDbCommand(updatestring, connection);
                    ucom.ExecuteNonQuery();
                    connection.Close();
                }
                catch
                {
                    MessageBox.Show("连接数据库失败！");
                    return;
                }

                MessageBox.Show("删除完毕！");
            }


        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string goal = "";
            try
            {
                goal = listView2.SelectedItems[0].Text ;
            }
            catch
            {
                MessageBox.Show("选定一个项目先！");
                return;
            }

            if ("Yes" == MessageBox.Show("重置会从今天开始重新计算，确定继续？", "??", MessageBoxButtons.YesNo).ToString())
            {
                string dayvalue = string.Format("{0:D}", DateTime.Now.Date);
                StartDate.Text = dayvalue;
                numericUpDown1.Value = 1;
                numericUpDown2.Value = 3;
                numericUpDown3.Value = 7;
                numericUpDown4.Value = 14;

                dateTimePicker1.Value = DateTime.Now.Date.AddDays(1);
                dateTimePicker2.Value = DateTime.Now.Date.AddDays(3);
                dateTimePicker3.Value = DateTime.Now.Date.AddDays(7);
                dateTimePicker4.Value = DateTime.Now.Date.AddDays(14);

                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Checked = true;

            }
        }
    }
}
