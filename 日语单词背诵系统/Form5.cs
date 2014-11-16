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
using DotNetSpeech;


namespace 日语单词背诵系统
{
    public partial class Form5 : Form
    {
        ArrayList selectedDates = new ArrayList();//想要在各窗体之间传递数据，定义static全局变量即可
        DataTable Words = new DataTable();
        int index = 0;
        OleDbDataAdapter myAda = new OleDbDataAdapter();
        DataSet myDs = new DataSet();

        
        bool c1 = true;
        bool c2 = true;
        bool c3 = true;
        bool c4 = true;
        bool c5 = true;
        bool c6 = true;
        bool c7 = true;
        bool c8 = true;
        bool c9 = true;

        bool d1 = true;
        bool d2 = true;
        bool d3 = true;
        bool d4 = true;
        bool d5 = true;
        bool d6 = true;
        bool d7 = true;
        bool d8 = true;
        bool d9 = true;

        bool display = false;
        public ArrayList goaldays = new ArrayList();
        string nowday = DateTime.Now.Date.ToString("yyyy-MM-dd");

        string dancibiao = form1.G_danci;
        string jihuabiao = form1.G_jihua;
        string voiced = form1.voice;

        public Form5()
        {
            InitializeComponent();
            selectedDates = Form3.selectedDates;
            toolTip.AutoPopDelay = 10000;
            toolTip.ToolTipTitle = "~~~~~~~~~~~~" + '\n' + '\n';

            toolTip.BackColor = Color.FromArgb(65, 65, 65);
            toolTip.ForeColor = Color.FromKnownColor(KnownColor.ButtonFace);
            
            toolTip.UseFading = true;

            

        }

        protected override void OnClosed(EventArgs e)
        {

            //Error:不返回任何键列信息的select语句，不支持commandbuilder的动态生成
            //此时，在数据库表中定义主键，而且在select语句中加入主键即可

            OleDbCommandBuilder db = new OleDbCommandBuilder(myAda);
            myAda.Update(Words);
            //myDs.AcceptChanges();  //另一种更新数据库的方法，不会

            base.OnClosed(e);

        }

        public void DisDone()
        {
            goaldays.Clear();
            display = false;

            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();
            string sqlstring = "select 描述 from "+jihuabiao +" where 第一次日期 = '" + nowday + "' or  第二次日期 = '" + nowday + "' or  第三次日期 = '" + nowday + "' or  第四次日期 = '" + nowday + "' ;";
            try
            {
                OleDbCommand myCommand = new OleDbCommand(sqlstring, connection);
                OleDbDataReader myReader = myCommand.ExecuteReader();
                while (myReader.Read())
                {
                    if (!goaldays.Contains(myReader.GetString(0)))
                    {
                        if (selectedDates.Contains(myReader.GetString(0)))
                        {
                            goaldays.Add(myReader.GetString(0));
                        }
                    }
                }
                myReader.Close();
                connection.Close();
            }
            catch
            {
                MessageBox.Show("发生错误", "提示！", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);

            }
            if (goaldays.Count == 0)
            {
                display = false;
            }
            else
            {
                display = true;
            }
        }

        public void StartConn(string str)
        {
            index = 0;
            string sqlColumn = "select 序号,平假名,汉字,释义,例句,词性,单词级别,总背诵次数,背诵重要级别,背诵标志,描述 from " + dancibiao + " where (描述 =";
            string sqlDate = "";
            string sqlOrder = str;
            string sqlString = "";
            try
            {
                //加载多个日期段的组
                if (Form3.selectedDates.Count == 0)
                {
                    MessageBox.Show("没有选择单词组！");
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
                        sqlDate += ("or  描述 ='" + Form3.selectedDates[i].ToString() + "'");
                    }
                    sqlDate = sqlDate + ")";
                }
                //


                //sqlOrder = " ORDER BY right(cstr(rnd(-int(rnd(-timer())*100+序号)))*1000*Now(),2) ";

                sqlString = sqlColumn + sqlDate + sqlOrder + ";";
                if (checkBox_bj2.Checked)
                {
                    sqlString = sqlColumn + sqlDate + "and 背诵标志=true" + sqlOrder + ";";
                }
                
                Console.WriteLine(sqlString);
            }
            catch
            {
                MessageBox.Show("先选择要背诵的单词组！");
                return;
            }
            try
            {
                Conn conn = new Conn();
                OleDbConnection connection = conn.CreatConn();
                connection.Open();
                myAda = new OleDbDataAdapter(sqlString, connection);


                //myAda.Fill(Words);

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
        }

        public string ContentEdit(int i, int j, string ss)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (ss == "平假名")
                {
                    content = Words.Rows[i + j]["平假名"].ToString();
                    cts = content.Split('【');
                    if (form1.Yinbiao)
                    {
                        content = cts[0].Trim ();
                    }
                    else
                    {
                        content = cts[0].Trim() + '\n' + '【' + cts[1].Trim();
                    }
                    
                }
                else if (ss == "汉字")
                {
                    content = Words.Rows[i + j]["汉字"].ToString().Trim();
                }
                else if (ss == "释义")
                {
                    //content = "<html><font color =red>FFFFFFFFFFuck</font></html>";
                    content = Words.Rows[i + j]["释义"].ToString().Trim();
                }
                else
                {
                    content = "fakenda";
                }
            }
            catch
            {

            }
            return content;
        }

        public string ContentEdit2(int i, int j, string ss)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (ss == "平假名")
                {
                    content = Words.Rows[i + j]["平假名"].ToString();
                    cts = content.Split('【');
                    if (cts .Length <2)
                    {
                        content = "         "+cts[0].Trim();
                    }
                    else
                    {
                        int kk = cts[0].Trim().Length;
                        string str = "";
                        for (int n = 0; n < 20 - 2*kk; n++)
                        {
                            str += ' ';
                        }
                        content = "         " + cts[0].Trim() + str + '【' + cts[1].Trim();
                    }

                }
                else if (ss == "汉字")
                {
                    content = "         " + Words.Rows[i + j]["汉字"].ToString().Trim();
                }
                else if (ss == "释义")
                {
                    //content = "<html><font color =red>FFFFFFFFFFuck</font></html>";
                    content = Words.Rows[i + j]["释义"].ToString().Trim();
                }
                else
                {
                    content = "fakenda";
                }
            }
            catch
            {

            }
            return content;
        }

        public void TT_show(int i)
        {
            string content = "";
            if (checkBox_hz.Checked)
            {
                content = ContentEdit(i, 0, "汉字");
                bb1.Text = content;
                content = ContentEdit(i, 1, "汉字");
                bb2.Text = content;
                content = ContentEdit(i, 2, "汉字");
                bb3.Text = content;
                content = ContentEdit(i, 3, "汉字");
                bb4.Text = content;
                content = ContentEdit(i, 4, "汉字");
                bb5.Text = content;
                content = ContentEdit(i, 5, "汉字");
                bb6.Text = content;
                content = ContentEdit(i, 6, "汉字");
                bb7.Text = content;
                content = ContentEdit(i, 7, "汉字");
                bb8.Text = content;
                content = ContentEdit(i, 8, "汉字");
                bb9.Text = content;

                c1 = true;
                c2 = true;
                c3 = true;
                c4 = true;
                c5 = true;
                c6 = true;
                c7 = true;
                c8 = true;
                c9 = true;

            }
            else
            {
                content = ContentEdit(i, 0, "平假名");
                bb1.Text = content;
                content = ContentEdit(i, 1, "平假名");
                bb2.Text = content;
                content = ContentEdit(i, 2, "平假名");
                bb3.Text = content;
                content = ContentEdit(i, 3, "平假名");
                bb4.Text = content;
                content = ContentEdit(i, 4, "平假名");
                bb5.Text = content;
                content = ContentEdit(i, 5, "平假名");
                bb6.Text = content;
                content = ContentEdit(i, 6, "平假名");
                bb7.Text = content;
                content = ContentEdit(i, 7, "平假名");
                bb8.Text = content;
                content = ContentEdit(i, 8, "平假名");
                bb9.Text = content;

                c1 = false;
                c2 = false;
                c3 = false;
                c4 = false;
                c5 = false;
                c6 = false;
                c7 = false;
                c8 = false;
                c9 = false;
            }



            label1.Text = "第 " + (index / 9 + 1) + " 页";

        }

        public void TT_show2(int i)
        {
            string content = "";
            if (checkBox_hz.Checked)
            {
                content = ContentEdit2(i, 0, "汉字");
                cc1.Text = content;
                content = ContentEdit2(i, 1, "汉字");
                cc2.Text = content;
                content = ContentEdit2(i, 2, "汉字");
                cc3.Text = content;
                content = ContentEdit2(i, 3, "汉字");
                cc4.Text = content;
                content = ContentEdit2(i, 4, "汉字");
                cc5.Text = content;
                content = ContentEdit2(i, 5, "汉字");
                cc6.Text = content;
                content = ContentEdit2(i, 6, "汉字");
                cc7.Text = content;
                content = ContentEdit2(i, 7, "汉字");
                cc8.Text = content;
                content = ContentEdit2(i, 8, "汉字");
                cc9.Text = content;

                d1 = true;
                d2 = true;
                d3 = true;
                d4 = true;
                d5 = true;
                d6 = true;
                d7 = true;
                d8 = true;
                d9 = true;


            }
            else
            {
                content = ContentEdit2(i, 0, "平假名");
                cc1.Text = content;
                content = ContentEdit2(i, 1, "平假名");
                cc2.Text = content;
                content = ContentEdit2(i, 2, "平假名");
                cc3.Text = content;
                content = ContentEdit2(i, 3, "平假名");
                cc4.Text = content;
                content = ContentEdit2(i, 4, "平假名");
                cc5.Text = content;
                content = ContentEdit2(i, 5, "平假名");
                cc6.Text = content;
                content = ContentEdit2(i, 6, "平假名");
                cc7.Text = content;
                content = ContentEdit2(i, 7, "平假名");
                cc8.Text = content;
                content = ContentEdit2(i, 8, "平假名");
                cc9.Text = content;

                d1 = false;
                d2 = false;
                d3 = false;
                d4 = false;
                d5 = false;
                d6 = false;
                d7 = false;
                d8 = false;
                d9 = false;

                
            }

            labelItem1.Text = TT_detailShow(i, 0);
            labelItem2.Text = TT_detailShow(i, 1);
            labelItem3.Text = TT_detailShow(i, 2);
            labelItem4.Text = TT_detailShow(i, 3);
            labelItem5.Text = TT_detailShow(i, 4);
            labelItem6.Text = TT_detailShow(i, 5);
            labelItem7.Text = TT_detailShow(i, 6);
            labelItem8.Text = TT_detailShow(i, 7);
            labelItem9.Text = TT_detailShow(i, 8);



            sideBar1.Refresh();

            label1.Text = "第 " + (index / 9 + 1) + " 页";

        }

        public string  TT_detailShow(int i,int j)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (!checkBox_hz.Checked)
                {
                    content = "     " + Words.Rows[i + j]["汉字"].ToString().Trim() + '\n' + '\n' + Words.Rows[i + j]["释义"].ToString().Trim() + '\n' + '\n' + Words.Rows[i + j]["例句"].ToString().Trim();
                }
                else
                {
                    content = "     " + Words.Rows[i + j]["平假名"].ToString().Trim() + '\n' + '\n' + Words.Rows[i + j]["释义"].ToString().Trim() + '\n' + '\n' + Words.Rows[i + j]["例句"].ToString().Trim();
                }
            }
            catch
            {

            }
            return content;
                    
        }

        public void Read(int i)
        {
            SpVoiceClass voice = new SpVoiceClass();
            try
            {
                
                voice.Voice = voice.GetVoices("name=" + voiced, string.Empty).Item(0);
                //voice.AllowAudioOutputFormatChangesOnNextSet = true;
                voice.Rate = -5;
                string rtext = Words.Rows[i]["平假名"].ToString ();
                rtext = rtext.Split('【')[0].ToString();
                try
                {
                    voice.Speak(rtext, SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak);
                }
                catch
                {
                    MessageBox.Show("未注册的语音包!请重新安装!");
                }
                
            }
            catch
            {
                MessageBox.Show("语音插件没有安装或正确配置，不能使用！");
                return;
            }
            voice = null;
            GC.Collect();
        }

        public void Mark(int i)
        {
            Words.Rows[i]["背诵标志"]=true ;
        }



        public void bb1_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c1)
                {
                    content = Words.Rows[i]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb1.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i]["汉字"].ToString();
                    //bb1.ForeColor = Color.FromArgb(192, 64, 0);
                }
            }
            catch
            {

            }
            bb1.Text = content;

            c1 = !c1;
        }

        public void bb2_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c2)
                {
                    content = Words.Rows[i + 1]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb2.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i + 1]["汉字"].ToString();
                    //bb2.ForeColor = Color.FromArgb(192, 64, 0);
                }
            }
            catch
            {

            }
            bb2.Text = content;

            c2 = !c2;
        }

        public void bb3_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c3)
                {
                    content = Words.Rows[i + 2]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb3.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i + 2]["汉字"].ToString();
                    //bb3.ForeColor = Color.FromArgb(192, 64, 0);
                }
            }
            catch
            {

            }
            bb3.Text = content;

            c3 = !c3;
        }

        public void bb4_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c4)
                {
                    content = Words.Rows[i + 3]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb4.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i + 3]["汉字"].ToString();
                    //bb4.ForeColor = Color.FromArgb(192, 64, 0);
                }
            }
            catch
            {

            }
            bb4.Text = content;

            c4 = !c4;
        }

        public void bb5_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c5)
                {
                    content = Words.Rows[i + 4]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb5.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i + 4]["汉字"].ToString();
                    //bb5.ForeColor = Color.FromArgb(192, 64, 0);
                }

            }
            catch
            {

            }
            bb5.Text = content;

            c5 = !c5;
        }

        public void bb6_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c6)
                {
                    content = Words.Rows[i + 5]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb6.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i + 5]["汉字"].ToString();
                    //bb6.ForeColor = Color.FromArgb(192, 64, 0);
                }
            }
            catch
            {

            }
            bb6.Text = content;

            c6 = !c6;
        }

        public void bb7_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c7)
                {
                    content = Words.Rows[i + 6]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb7.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i + 6]["汉字"].ToString();
                    //bb7.ForeColor = Color.FromArgb(192, 64, 0);
                }
            }
            catch
            {

            }
            
            bb7.Text = content;

            c7 = !c7;
        }

        public void bb8_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c8)
                {
                    content = Words.Rows[i + 7]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb8.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i + 7]["汉字"].ToString();
                    //bb8.ForeColor = Color.FromArgb(192, 64, 0);
                }

            }
            catch
            {

            }
            
            bb8.Text = content;

            c8 = !c8;
        }

        public void bb9_show(int i)
        {
            string content = "";
            string[] cts = new string[2];
            try
            {
                if (c9)
                {	

                    content = Words.Rows[i + 8]["平假名"].ToString();
                    cts = content.Split('【');
                    content = cts[0] + '\n' + '【' + cts[1];
                    //bb9.ForeColor = Color.FromArgb(192, 0, 0);
                }
                else
                {
                    content = Words.Rows[i + 8]["汉字"].ToString();
                    //bb9.ForeColor = Color.FromArgb(192, 64, 0);
                }
            }
            catch
            {

            }
            
            bb9.Text = content;

            c9 = !c9;
        }






        private void bb1_Click_1(object sender, EventArgs e)
        {
            bb1_show(index);
        }

        private void bb2_Click(object sender, EventArgs e)
        {
            bb2_show(index);
        }

        private void bb3_Click(object sender, EventArgs e)
        {
            bb3_show(index);
        }

        private void bb4_Click(object sender, EventArgs e)
        {
            bb4_show(index);
        }

        private void bb5_Click(object sender, EventArgs e)
        {
            bb5_show(index);
        }

        private void bb6_Click(object sender, EventArgs e)
        {
            bb6_show(index);
        }

        private void bb7_Click(object sender, EventArgs e)
        {
            bb7_show(index);
        }

        private void bb8_Click(object sender, EventArgs e)
        {
            bb8_show(index);
        }

        private void bb9_Click(object sender, EventArgs e)
        {
            bb9_show(index);
        }

        private void button_START_Click(object sender, EventArgs e)
        {

            Words.Rows.Clear();
            string sqlOrder = "";
            if (radioButton1.Checked)
            {
                sqlOrder = " ORDER BY right(cstr(rnd(-int(rnd(-timer())*100+序号)))*1000*Now(),2) ";
            }
            else if (radioButton2.Checked)
            {

            }
            else
            {
                sqlOrder = "  ORDER BY 背诵重要级别 desc";
            }
            StartConn(sqlOrder);
            TT_show(index);
            TT_show2(index);

            trackBar1.SetRange(1, Words.Rows.Count);//设置trackbar的范围
            trackBar1.TickFrequency = 9; //设置trackbar控件显示的刻度最小值
            trackBar1.SmallChange = 1;  //设置trackbar每次改变的最小值
            trackBar1.LargeChange = 9;  //设置trackbar每次改变的最大值

            label1.Text = "第 1 页 ";
            int kn = Words.Rows.Count / 9 + 1;
            label2.Text = "/ 共  " + kn + " 页";

            DisDone();
            if (display)
            {
                this.button_Done.Visible = true;
            }
        }

        private void button_CHOICE_Click(object sender, EventArgs e)
        {
            Dialog2 dialog = new Dialog2();
            dialog.ShowDialog();
        }
        
        private void button_down_Click(object sender, EventArgs e)
        {
            if (index + 9 <= Words.Rows.Count)
            {
                index += 9;
            }
            TT_show(index);
            TT_show2(index);
            
        }

        private void button_up_Click(object sender, EventArgs e)
        {
            if (index - 9 >= 0)
            {
                index -= 9;
            }
            TT_show(index);
            TT_show2(index);
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            index = trackBar1.Value - 1;
            TT_show(index);
            TT_show2(index);
        }


        //private void MouseEnter(object sender, EventArgs e)
        //{
        //    int i = 0;
        //    Button bb=sender as Button ;
        //    if (bb==null )
        //    {
        //        return;
        //    }
        //    else 
        //    {
        //        i = int.Parse(bb.Tag);
        //    }
            
        //}

        private void bb1_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 0, "释义");
            if (c1)
            {
                toolTip.Show(content, bb1, 150, 100);
            }
        }

        private void bb1_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb1);
        }

        private void bb2_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 1, "释义");
            if (c2)
            {
                toolTip.Show(content, bb2, 150, 100);
            }
        }

        private void bb2_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb2);
        }

        private void bb3_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 2, "释义");
            if (c3)
            {
                toolTip.Show(content, bb3, 150, 100);
            }
        }

        private void bb3_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb3);
        }

        private void bb4_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 3, "释义");
            if (c4)
            {
                toolTip.Show(content, bb4, 150, 100);
            }
        }

        private void bb4_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb4);
        }

        private void bb5_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 4, "释义");
            if (c5)
            {
                toolTip.Show(content, bb5, 150, 100);
            }
        }

        private void bb5_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb5);
        }

        private void bb6_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 5, "释义");
            if (c6)
            {
                toolTip.Show(content, bb6, 150, 100);
            }
        }

        private void bb6_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb6);
        }

        private void bb7_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 6, "释义");
            if (c7)
            {
                toolTip.Show(content, bb7, 150, 100);
            }
        }

        private void bb7_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb7);
        }

        private void bb8_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 7, "释义");
            if (c8)
            {
                toolTip.Show(content, bb8, 150, 100);
            }
        }

        private void bb8_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb8);
        }

        private void bb9_MouseEnter(object sender, EventArgs e)
        {
            string content = ContentEdit(index, 8, "释义");
            if (c9)
            {
                toolTip.Show(content, bb9, 150, 100);
            }
        }

        private void bb9_MouseLeave(object sender, EventArgs e)
        {
            toolTip.Hide(bb9);
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)//object sender、Tag的用途
        {
            int i = 0;
            i = int.Parse(contextMenuStrip1.SourceControl.Tag.ToString());

            Read(index+i);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)//object sender、Tag的用途
        {
            int i = 0;
            i = int.Parse(contextMenuStrip1.SourceControl.Tag.ToString());

            Mark(index + i);
        }

        private void button_Done_Click(object sender, EventArgs e)
        {

            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();


            string upstr = "";

            string sql1 = "select 第一次日期,第二次日期,第三次日期,第四次日期, 描述 from " + jihuabiao + " where 描述 ='";
            foreach (string date in selectedDates)
            {
                string sql2 = date + "'";
                string sqlstr = sql1 + sql2;
                OleDbCommand myCommand = new OleDbCommand(sqlstr, connection);
                OleDbDataReader myReader;
                try
                {
                    myReader = myCommand.ExecuteReader();
                }
                catch
                {
                    break;
                }

                try
                {
                    while (myReader.Read())
                    {
                        if (nowday == myReader.GetString(0))
                        {
                            upstr = "update " + jihuabiao + " set 第一次完成情况 = true where 描述 ='" + date + "'";
                        }
                        else if (nowday == myReader.GetString(1))
                        {
                            upstr = "update " + jihuabiao + " set 第二次完成情况 = true where 描述 ='" + date + "'";
                        }
                        else if (nowday == myReader.GetString(2))
                        {
                            upstr = "update " + jihuabiao + " set 第三次完成情况 = true where 描述 ='" + date + "'";
                        }
                        else if (nowday == myReader.GetString(3))
                        {
                            upstr = "update " + jihuabiao + " set 第四次完成情况 = true where 描述 ='" + date + "'";
                        }

                        myCommand.CommandText = upstr;
                        myReader.Close();
                        myCommand.ExecuteNonQuery();


                        //不可以这样，最后UPdate（）时会发生错误
                        //string sqladd = "UPDATE 单词总表 SET 总背诵次数 = 总背诵次数+1 where 录入日期 ='" + sql2;
                        //myCommand.CommandText = sqladd;
                        //myCommand.ExecuteNonQuery();

                    }
                }
                catch
                {
                    //break;
                }

            }
            connection.Close();
            try
            {
                foreach (string date2 in selectedDates)
                {
                    for (int i = 0; i < Words.Rows.Count; i++)
                    {
                        if (Words.Rows[i]["描述"].ToString() == date2)
                        {
                            string k = (int.Parse(Words.Rows[i]["总背诵次数"].ToString()) + 1).ToString();
                            Words.Rows[i]["总背诵次数"] = k;
                        }
                    }
                }
            }
            catch
            {

            }
            MessageBox.Show("WELL DONE!");
            Close();
        }




    }
}
