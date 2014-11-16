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
    
    public partial class Form3 : Form
    {
        public static ArrayList selectedDates = new ArrayList();//想要在各窗体之间传递数据，定义static全局变量即可
        //为何还得用new?

        //bool edit = true ;
        bool info = false;
        bool jm = false;
        bool hz = false;
        bool sy = false;
        bool lj = false;
        public DataTable Words = new DataTable();
        int index = 0;
        OleDbDataAdapter myAda = new OleDbDataAdapter();
        DataSet myDs = new DataSet();

        bool display = false;
        public ArrayList goaldays = new ArrayList();
        string nowday = DateTime.Now.Date.ToString("yyyy-MM-dd");

        string dancibiao = form1.G_danci;
        string jihuabiao = form1.G_jihua;
        string voiced = form1.voice;

        public Form3()
        {
            InitializeComponent();
            contextMenuStrip1.BackColor = Color.FromArgb(65, 65, 65);
            contextMenuStrip1.ForeColor = Color.FromKnownColor(KnownColor.ButtonFace);
        }


        protected override void  OnClosed(EventArgs e)
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
            goaldays .Clear ();
            display = false;

            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();

            string sqlstring = "select 描述 from " + jihuabiao + " where 第一次日期 = '" + nowday + "' or  第二次日期 = '" + nowday + "' or  第三次日期 = '" + nowday + "' or  第四次日期 = '" + nowday + "' ;";
            try
            {
                OleDbCommand myCommand = new OleDbCommand(sqlstring, connection);
                OleDbDataReader myReader = myCommand.ExecuteReader();
                while (myReader.Read())
                {
                    if (!goaldays .Contains(myReader.GetString(0)))
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

        public void Read(int i,string s)
        {
            SpVoiceClass voice = new SpVoiceClass();
            string rtext = "";
            string[] cts = new string[10];
            try
            {

                voice.Voice = voice.GetVoices("name="+voiced, string.Empty).Item(0);
                
                //voice.AllowAudioOutputFormatChangesOnNextSet = true;
                voice.Rate = -7;
            }
            catch
            {
                MessageBox.Show("语音插件没有安装或正确配置，不能使用！");
                return;
            }

            try
            {
                if (s == "平假名")
                {
                    rtext = Words.Rows[i]["平假名"].ToString();
                    cts = rtext.Split('【');
                    rtext = cts[0];
                }
                else if (s == "例句")
                {
                    rtext = Words.Rows[i]["例句"].ToString();
                    cts = rtext.Split('／');
                    rtext = cts[0];
                }
                else
                {
                    rtext = "危険信号";
                }
            }
            catch
            {
                rtext = "危険信号";
            }
            try
            {
                voice.Speak(rtext, SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak);
            }
            catch
            {
                MessageBox .Show ("未注册的语音包!请重新安装!");
            }

            voice = null;
            GC.Collect();

        }


        public void INFO_show(int i)
        {
            if (Words.Rows.Count == 0)
            {
                return;
            }
            info = !info;
            Graphics g = groupBox1.CreateGraphics();
            string content = Words.Rows[i]["词性"].ToString()+"  "+Words.Rows[i]["单词级别"].ToString()+"级词汇";
            SolidBrush sbrush = new SolidBrush(Color.FromArgb(0, 192, 0));
            SolidBrush ebrush = new SolidBrush(SystemColors.Control);

            try
            {
                if (info)
                {
                    g.DrawString(content, new Font("微软雅黑", 14), sbrush, 200, 84);
                }
                else
                {
                    g.DrawString(content, new Font("微软雅黑", 14), ebrush, 200, 84);
                    Rectangle rc = new Rectangle(200, 84, 200, 50);
                    groupBox1.Invalidate(rc, true);
                    groupBox1.Update();
                   
                }
            }
            finally
            {
                
                g.Dispose();
            }

        }

        public void JM_show(int i)
        {
            if (Words.Rows.Count == 0)
            {
                return;
            }
            jm = !jm;
            Graphics g = groupBox1.CreateGraphics();
            string content = Words.Rows[i]["平假名"].ToString();
            SolidBrush sbrush = new SolidBrush(Color.FromArgb(192, 0, 0));
            SolidBrush ebrush = new SolidBrush(SystemColors .Control );
            
            try
            {
                if (jm)
                {
                    g.DrawString(content, new Font("微软雅黑", 14), sbrush, 200, 137);
                }
                else
                {
                    g.DrawString(content, new Font("微软雅黑", 14), ebrush, 200, 137);
                    Rectangle rc = new Rectangle(200, 137, 200, 50);
                    groupBox1.Invalidate(rc, true);
                    groupBox1.Update();
                }
            }
            finally
            {
                g.Dispose();
            }
            
        }

        public void HZ_show(int i)
        {
            if (Words.Rows.Count == 0)
            {
                return;
            }
            hz = !hz;
            Graphics g = groupBox1.CreateGraphics();
            string content = Words.Rows[i]["汉字"].ToString();
            SolidBrush sbrush = new SolidBrush(Color.FromArgb(192, 64, 0));
            SolidBrush ebrush = new SolidBrush(SystemColors.Control);
            try
            {
                if (hz)
                {
                    g.DrawString(content, new Font("微软雅黑", 14), sbrush, 250, 190);
                }
                else
                {
                    g.DrawString(content, new Font("微软雅黑", 14), ebrush, 250, 190);
                    Rectangle rc = new Rectangle(250, 190, 200, 50);
                    groupBox1.Invalidate(rc, true);
                    groupBox1.Update();
                }
            }
            finally
            {
                g.Dispose();
            }
        }

        public void SY_show(int i)
        {
            if (Words.Rows.Count == 0)
            {
                return;
            }
            sy = !sy;
            Graphics g = groupBox1.CreateGraphics();
            string content = Words.Rows[i]["释义"].ToString();
            SolidBrush sbrush = new SolidBrush(Color.FromArgb(192, 192, 0));
            SolidBrush ebrush = new SolidBrush(SystemColors.Control);
            Rectangle rc = new Rectangle(190, 240, 307, 80);
            try
            {
                if (sy)
                {
                    g.DrawString(content, new Font("微软雅黑", 14), sbrush, rc);
                }
                else
                {
                    g.DrawString(content, new Font("微软雅黑", 14), ebrush, rc);
                   
                    groupBox1.Invalidate(rc, true);
                    groupBox1.Update();
                }
            }
            finally
            {
                g.Dispose();
            }
        }

        public void LJ_show(int i)
        {
            if (Words.Rows.Count == 0)
            {
                return;
            }
            lj = !lj;
            Graphics g = groupBox1.CreateGraphics();
            string content = Words.Rows[i]["例句"].ToString();
            SolidBrush sbrush = new SolidBrush(Color.FromArgb(0,192, 192));
            SolidBrush ebrush = new SolidBrush(SystemColors.Control);
            Rectangle rc = new Rectangle(180, 330, 310, 110);
            try
            {
                if (lj)
                {
                    g.DrawString(content, new Font("微软雅黑", 12), sbrush,rc);
                }
                else
                {
                    g.DrawString(content, new Font("微软雅黑", 12), ebrush,rc);
                    
                    groupBox1.Invalidate(rc, true);
                    groupBox1.Update();
                }
            }
            finally
            {
                g.Dispose();
            }
        }

        public void TT_show(int i)
        {
            //this.Invalidate(true);//请求windows为窗体和子控件触发一个paint事件
            //this.Update();//强迫并立即触发paingt事件

            if (Words.Rows.Count == 0)
            {
                return;
            }
            info = false ;
            jm = false;
            hz = false;
            sy = false;
            lj = false;


            Graphics g = groupBox1.CreateGraphics();
            g.Clear(SystemColors.Control);  //清空背景

            if (checkBox_info .Checked)
            {
                string content = Words.Rows[i]["词性"].ToString() + "  " + Words.Rows[i]["单词级别"].ToString() + "级词汇";
                SolidBrush sbrush = new SolidBrush(Color.FromArgb(0, 192, 0));
                g.DrawString(content, new Font("微软雅黑", 14), sbrush, 200, 84);
                info = !info;
            }
            
            if (checkBox_jm.Checked)
            {
                string content = Words.Rows[i]["平假名"].ToString();
                SolidBrush sbrush = new SolidBrush(Color.FromArgb(192, 0, 0));
                g.DrawString(content, new Font("微软雅黑", 14), sbrush, 200, 137);
                jm = !jm;
            }

            if (checkBox_hz.Checked)
            {
                string content = Words.Rows[i]["汉字"].ToString();
                SolidBrush sbrush = new SolidBrush(Color.FromArgb(192, 64, 0));
                g.DrawString(content, new Font("微软雅黑", 14), sbrush, 250, 190);
                hz = !hz;
            }

            if (checkBox_sy.Checked)
            {
                string content = Words.Rows[i]["释义"].ToString();
                SolidBrush sbrush = new SolidBrush(Color.FromArgb(192, 192, 0));
                Rectangle rc = new Rectangle(190, 240, 307, 80);
                g.DrawString(content, new Font("微软雅黑", 14), sbrush, rc);//用rc来Drawstring可以实现自动换行
                sy = !sy;
            }
            if (checkBox_lj.Checked)
            {
                string content = Words.Rows[i]["例句"].ToString();
                SolidBrush sbrush = new SolidBrush(Color.FromArgb(0, 192, 192));
                SolidBrush ebrush = new SolidBrush(SystemColors.Control);
                Rectangle rc = new Rectangle(180, 330, 310, 110);
                g.DrawString(content, new Font("微软雅黑", 12), sbrush, rc);
                lj = !lj;
            }
            g.Dispose();

            textBox_nandu.Text = Words.Rows[i]["背诵重要级别"].ToString();
            checkBox_fxbz.Checked = (bool )Words.Rows[i]["背诵标志"];


            
        }


        private void button11_Click(object sender, EventArgs e)
        {
            Words.Rows.Clear();
            if (checkBox_nd.Checked == false && checkBox_sx.Checked == false && checkBox_sj.Checked == false)
            {
                MessageBox.Show("先选择要背诵排序方式！");
                return;
            }
            index = 0;

            string sqlColumn = "select 序号,平假名,汉字,释义,例句,词性,单词级别,录入日期,总背诵次数,背诵重要级别,背诵标志,描述 from " + dancibiao + " where (描述 =";//where (录入日期 =
            string sqlDate = "";
            string sqlOrder = "";
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
                        sqlDate += ("or  描述 ='" + Form3.selectedDates[i].ToString() + "'");//"or  录入日期 ='"
                    }
                    sqlDate = sqlDate + ")";
                }
                //


                //根据不同的排序选项进行选择不同的语句
                if (checkBox_sj.Checked)
                {
                    sqlOrder = " ORDER BY right(cstr(rnd(-int(rnd(-timer())*100+序号)))*1000*Now(),2) ";
                    //随机排序函数。好东西！
                }

                else if (checkBox_nd.Checked)
                {
                    sqlOrder = "  ORDER BY 背诵重要级别 desc";
                }

                else
                {

                }
                //
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


                myAda.Fill(Words);

                //myAda.Fill(myDs, "myWords");
                //Words = myDs.Tables["myWords"];

                //DataTable myWords = new DataTable();
                //myAda.Fill(myWords);//不用myDs，直接导入到表中即可




                connection.Close();


            }
            catch
            {
                MessageBox.Show("检查数据库联接！");
                return;
            }



            label1.Text = "1";
            label2 .Text ="/    "+Words.Rows.Count.ToString();


            trackBar1.SetRange(1, Words.Rows.Count);//设置trackbar的范围
            trackBar1.TickFrequency = 5; //设置trackbar控件显示的刻度最小值
            trackBar1.SmallChange = 1;  //设置trackbar每次改变的最小值
            trackBar1.LargeChange = 5;  //设置trackbar每次改变的最大值

            DisDone();
            if (display)
            {
                this.button_Done.Visible = true;
            }

            TT_show(index);


            

        }



        private void button10_Click(object sender, EventArgs e)
        {
            selectedDates.Clear();
            Dialog2 form5 = new Dialog2();
            form5.ShowDialog();
        }

        private void button_jm_Click(object sender, EventArgs e)
        {
            JM_show(index);
        }

        private void button_hz_Click(object sender, EventArgs e)
        {
            HZ_show(index);
        }

        private void button_sy_Click(object sender, EventArgs e)
        {
            SY_show(index);
        }

        private void button_lj_Click(object sender, EventArgs e)
        {
            LJ_show(index);
        }

        private void button_info_Click(object sender, EventArgs e)
        {
            INFO_show(index);
        }

        private void button_down_Click(object sender, EventArgs e)
        {
            if (index + 1 <= Words.Rows.Count-1)
            {
                index += 1;
                TT_show(index);
                int k=index+1;
                label1.Text = k.ToString ();
            }
            else
            {
                MessageBox.Show("已经是最后了！^_^");
            }
        }

        private void button_up_Click(object sender, EventArgs e)
        {
            if (index - 1 >= 0)
            {
                index -= 1;
                TT_show(index);
                int k = index + 1;
                label1.Text = k.ToString();
            }
            else
            {
                MessageBox.Show("已经是最前了！^_^");
            }
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            index = trackBar1.Value - 1;
            TT_show(index);
            int k = index + 1;
            label1.Text = k.ToString();

        }

        private void checkBox_sj_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_sj.Checked)
            {
                checkBox_sx.Checked = false;
                checkBox_nd.Checked = false;
            }
        }

        private void checkBox_sx_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_sx.Checked)
            {
                checkBox_sj.Checked = false;
                checkBox_nd.Checked = false;
            }
        }

        private void checkBox_nd_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_nd.Checked)
            {
                checkBox_sj.Checked = false;
                checkBox_sx.Checked = false;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (textBox_nandu.Text == "")
            {
                return;
            }
            textBox_nandu.Text=Convert .ToString(int.Parse (textBox_nandu .Text )+1);
            
        }

        private void textBox_nandu_TextChanged(object sender, EventArgs e)
        {
            Words.Rows[index]["背诵重要级别"] = textBox_nandu.Text;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox_nandu.Text == "")
            {
                return;
            }
            textBox_nandu.Text = Convert.ToString(int.Parse(textBox_nandu.Text) - 1);
        }

        private void checkBox_fxbz_CheckedChanged(object sender, EventArgs e)
        {
            if (Words.Rows.Count == 0)
            {
                return;
            }
            Words.Rows[index]["背诵标志"] = (bool )checkBox_fxbz.Checked;
        }

        private void 朗读假名ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Read(index, "平假名");
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Read(index, "例句");
        }


        private void button_Done_Click(object sender, EventArgs e)
        {
            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();


            string upstr = "";

            string sql1 = "select 第一次日期,第二次日期,第三次日期,第四次日期, 描述 from " + jihuabiao + " where 描述 ='";//第四次日期, 单词组
            foreach (string date in selectedDates)
            {
                string sql2 = date+"'";
                string sqlstr = sql1 + sql2;
                OleDbCommand myCommand = new OleDbCommand(sqlstr, connection);
                OleDbDataReader myReader ;
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
                        if (nowday==myReader.GetString(0))
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
