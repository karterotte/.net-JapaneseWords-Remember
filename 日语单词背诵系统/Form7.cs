using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;


namespace 日语单词背诵系统
{
    public partial class Form7 : Form
    {
        public DataTable Words = new DataTable();
        OleDbDataAdapter myAda = new OleDbDataAdapter();
        DataSet myDs = new DataSet();
        int total = 0;

        string dancibiao = form1.G_danci;
        string jihuabiao = form1.G_jihua;

        public Form7()
        {
            InitializeComponent();
            comboBox1.Text = form1.wt;
            comboBox2.Text = form1.wl;

        }
        public Table  DoTable(string jiaming,string hanzi,string shiyi,string liju,string cixing,string jibie)
        {

            Table table = new Table(3);
            table.CellsFitPage = true;
            table.TableFitsPage = true;
            table.Width = 95f;//设置table宽度
            table.BorderWidth = 3;
            //table.BorderColor = new iTextSharp.text.Color(255, 0, 0);
            table.Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.BOTTOM_BORDER;

            table.Padding = 2;//设置单元格边界和内容间的空白
            table.Spacing = 4;//设置单元格和表格边界间的空白
            //table.Border = iTextSharp.text.Rectangle.NO_BORDER;
            BaseFont YHei=null ;
            try
            {
                YHei = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\msyh.TTF", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                //加载系统中字体；
            }
            catch
            {
                MessageBox.Show("抱歉，需要装微软雅黑字体先!");
                return null;
            }
            
            iTextSharp.text.Font JM = new iTextSharp.text.Font(YHei, 16, 0, iTextSharp.text.Color.BLACK );
            iTextSharp.text.Font HZ = new iTextSharp.text.Font(YHei, 14f, 0, iTextSharp.text.Color.RED);
            iTextSharp.text.Font SY = new iTextSharp.text.Font(YHei, 12f, 2, iTextSharp.text.Color.GRAY );
            iTextSharp.text.Font LJ = new iTextSharp.text.Font(YHei, 12f, 2, iTextSharp.text.Color.DARK_GRAY  );
            iTextSharp.text.Font OT = new iTextSharp.text.Font(YHei, 12f, 0, iTextSharp.text.Color.LIGHT_GRAY );


            Cell cell = new Cell(new Paragraph (jiaming ,JM ));
            cell.Header = true;
            cell.BorderWidth = 1;
            cell.Border = iTextSharp.text.Rectangle.BOTTOM_BORDER;
            cell.Colspan = 2;
            table.AddCell(cell);

            cell = new Cell(new Paragraph( "            "+cixing + "  " + jibie, OT));
            cell.Header = true;
            cell.BorderWidth = 1;
            cell.Border = iTextSharp.text.Rectangle.BOTTOM_BORDER;
            cell.Colspan = 1;
            table.AddCell(cell,0,2);
            
            cell = new Cell(new Paragraph(hanzi, HZ ));//这样可以设置字体形式
            cell.Colspan = 1;
            cell.Rowspan = 1;

            //cell.BorderColor = new iTextSharp.text.Color(0, 0, 255);
            cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
            table.AddCell(cell, 2, 0);

            cell = new Cell(new Paragraph(shiyi  , SY  ));
            cell.Colspan = 1;//Colspan，即属于1个单元格列
            cell.Rowspan = 3;//Rowspan，即属于3个单元格行
            cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
            table.AddCell(cell, 3, 0);//是绝对行数和列数，包括前面的延伸好的单元格


            cell = new Cell(new Paragraph(liju, LJ));

            cell.Colspan = 2;
            cell.Rowspan = 4;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
            table.AddCell(cell, 2, 1);
            return table;
            

            
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

                if (checkedListBox.Items.Count == 0 && key == true)
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

        private void button1_Click(object sender, EventArgs e)
        {
            string dayvalue = DateTime.Now.Date.ToString("yyyy-MM-dd");
            //DateTime dayvalue = Convert.ToDateTime(dayvalue2);
            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();
            string sqlstring = "select 单词组 from "+jihuabiao +" where 第一次日期 = '" + dayvalue + "' or  第二次日期 = '" + dayvalue + "' or  第三次日期 = '" + dayvalue + "' or  第四次日期 = '" + dayvalue + "' ;";



            readQuery(sqlstring, connection, checkedListBox1,false );
            connection.Close();
            if (checkedListBox1.Items.Count == 0)
            {
                MessageBox.Show("今天没有计划任务!");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();
            string sqlstring = "select distinct 录入日期 from "+dancibiao +";";
            readQuery(sqlstring, connection, checkedListBox1,true );
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Words.Rows.Count == 0)
            {
                MessageBox.Show("先生成单词表！");
                return;
            }
            else
            {
                Document document = new Document();
                document.AddAuthor("fakenda");
                string filename = "单词表  " + DateTime.Now.ToShortDateString() + ".pdf";
                string path = Application.StartupPath + @"\Records\" + filename;
                try
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(path, FileMode.Create));
                    total = (int)numericUpDown1.Value - (int)numericUpDown2.Value + 1;
                    //document.SetMargins(1f, 1f, 2f, 2f);
                    //document.SetMarginMirroring(true);
                    //document.GetBottom(1);
                    HeaderFooter header = new HeaderFooter(new Phrase("OOXX     " + DateTime.Now.ToShortDateString() + "         Total    " + total.ToString() + "  words"), false);
                    document.Header = header;//Header.Footer 需要在document.Open()前设置
                    //document.Footer = new HeaderFooter(new Phrase("This is page:   " ),true );

                    document.SetPageSize(iTextSharp.text.PageSize.A4);
                    document.Open();


                    for (int i = (int)numericUpDown2.Value - 1; i < (int)numericUpDown1.Value; i++)
                    {
                        Table table = DoTable(Words.Rows[i]["平假名"].ToString(), Words.Rows[i]["汉字"].ToString(), Words.Rows[i]["释义"].ToString(), Words.Rows[i]["例句"].ToString(), Words.Rows[i]["词性"].ToString(), Words.Rows[i]["单词级别"].ToString());
                        if (table == null)
                        {
                            Close();
                            return;

                        }
                        //document.Add(table);

                        if (!writer.FitsPage(table))
                        {
                            //table.DeleteLastRow();
                            //i--;
                            //table.Offset = 0;
                            document.NewPage();
                            document.Add(table);

                            //table = getTable();
                        }
                        else
                        {
                            //table.Offset = 32;
                            document.Add(table);

                        }
                    }
                }

                catch (DocumentException de)
                {
                    Console.Error.WriteLine(de.Message);
                }
                catch (IOException ioe)
                {
                    Console.Error.WriteLine(ioe.Message);
                }
                catch
                {
                    MessageBox.Show("生成PDF错误！");
                }
                document.Close();
                try
                {
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    MessageBox.Show("需要装PDF浏览器！生成的PDF保存在程序目录Records中(OOXX)");
                }
                }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form3.selectedDates.Clear();
            string[] rqs = { };

            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                Console.WriteLine(checkedListBox1.CheckedItems[i].ToString());

                Form3.selectedDates.Add(checkedListBox1.CheckedItems[i].ToString());
                //调用全局变量，必须加前面的类
            }

            Words.Rows.Clear();
            string sqlColumn = "select 序号,平假名,汉字,释义,例句,词性,单词级别,总背诵次数,背诵重要级别,背诵标志 from "+dancibiao +" where (录入日期 =";
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
                        sqlDate += ("or  录入日期 ='" + Form3.selectedDates[i].ToString() + "'");
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

                if (checkBox_bj.Checked)
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
            total =Words.Rows.Count;
            numericUpDown1.Value = total;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDown1.Value > total)
            {
                MessageBox.Show("超出最大数目！");
                numericUpDown1.Value = total;
            }
            else if (numericUpDown1.Value <= 0)
            {
                MessageBox.Show("不可为负数！");
                numericUpDown1.Value = total;
            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDown2.Value > total)
            {
                MessageBox.Show("超出最大数目！");
                numericUpDown2.Value = 1;
            }
            else if (numericUpDown2.Value <= 0)
            {
                MessageBox.Show("不可为负数！");
                numericUpDown1.Value = 1;
            }
            else if (numericUpDown2.Value > numericUpDown1.Value)
            {
                MessageBox.Show("超出限定数目！");
                numericUpDown2.Value = 1;
            }

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

        private void button9_Click(object sender, EventArgs e)
        {
            if (Words.Rows.Count == 0)
            {
                MessageBox.Show("先生成单词表！");
                return;
            }
            else
            {
                Document document = new Document();
                document.AddAuthor("fakenda");
                string filename = "测验卷  " + DateTime.Now.ToShortDateString() + ".pdf";
                string path = Application.StartupPath + @"\Records\" + filename;
                try
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(path, FileMode.Create));
                    total = (int)numericUpDown1.Value - (int)numericUpDown2.Value + 1;
                    //document.SetMargins(1f, 1f, 2f, 2f);
                    //document.SetMarginMirroring(true);
                    //HeaderFooter header = new HeaderFooter(new Phrase("OOXX     " + DateTime.Now.ToShortDateString() + "         Total    " + total.ToString() + "  words"), false);
                    //document.Header = header;//Header.Footer 需要在document.Open()前设置
                    //document.Footer = header;
                    document.SetPageSize(iTextSharp.text.PageSize.A4);
                    document.Open();

                    Table datatable = new Table(5);
                    datatable.CellsFitPage = true;
                    datatable.Width = 100f;//设置表整体的宽度
                    datatable.Padding = 3;
                    datatable.Spacing = 3;
                    //datatable.setBorder(Rectangle.NO_BORDER);
                    float[] headerwidths = { 25,25,1,25,25 };//设置表中列的宽度
                    datatable.Widths = headerwidths;
                    //datatable.WidthPercentage = 100;

                    BaseFont YHei = null;
                    try
                    {
                        YHei = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\msyh.TTF", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        //加载系统中字体；
                    }
                    catch
                    {
                        MessageBox.Show("抱歉，需要装微软雅黑字体先!");
                        return;
                    }
                    iTextSharp.text.Font JM = new iTextSharp.text.Font(YHei, 14, 2, iTextSharp.text.Color.BLACK);
                    iTextSharp.text.Font TT = new iTextSharp.text.Font(YHei, 16, 0, iTextSharp.text.Color.GRAY );

                    Cell cell = new Cell(new Phrase("fakenda@hotmail.com", TT));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.Leading = 30;
                    cell.Colspan = 5;
                    cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                    cell.BackgroundColor = new iTextSharp.text.Color(0xC0, 0xC0, 0xC0);
                    datatable.AddCell(cell);
                    datatable.EndHeaders();

                    datatable.DefaultCellBorderWidth = 2;
                    datatable.DefaultHorizontalAlignment = 1;
                    datatable.DefaultRowspan = 1;
                    datatable.DefaultColspan = 1;
                    datatable.DefaultHorizontalAlignment = Element.ALIGN_LEFT ;

                    string key = "";

                    if (checkBox4.Checked ==true)
                    {
                        key = "汉字";
                    }
                    else
                    {
                        key = "平假名";
                    }

                    for (int i = (int)numericUpDown2.Value - 1; i < (int)numericUpDown1.Value; i=i+2)
                    {
                        cell = new Cell(new Paragraph("(" + (i + 1).ToString() + ") " + Words.Rows[i][key].ToString().Split('【')[0].Trim(), JM));//这样可以设置字体形式

                        datatable.AddCell(cell);

                        cell = new Cell(new Paragraph("", JM));
                        datatable.AddCell(cell);

                        cell = new Cell(new Paragraph("", JM));
                        cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                        cell.BackgroundColor = iTextSharp.text.Color.BLACK;
                        datatable.AddCell(cell);

                        try
                        {
                            cell = new Cell(new Paragraph("(" + (i+2).ToString() + ") " + Words.Rows[i + 1][key ].ToString().Split('【')[0].Trim(), JM));
                            datatable.AddCell(cell);

                            cell = new Cell(new Paragraph("", JM));
                            datatable.AddCell(cell);
                        }
                        catch
                        {
                            break;
                        }

                    }
                    document.Add(datatable);
                    
                    document.NewPage();
                    document.Add(new Phrase("ANSWER:", JM));

                    PdfContentByte cb = writer.DirectContent;
                    ColumnText ct = new ColumnText(cb);
                    
                    if (key == "平假名")
                    {
                        key = "汉字";
                    }
                    else
                    {
                        key = "平假名";
                    }
                    for (int i = (int)numericUpDown2.Value - 1; i < (int)numericUpDown1.Value; i++)
                    {
                        ct.AddText(new Paragraph("(" + (i + 1).ToString() + ") " + Words.Rows[i][key].ToString().Split('【')[0].Trim()+'\n', JM));
                    }
                    ct.Indent = 10;//缩进
                    int status = 0;
                    int column = 0;
                    float[] right = { 50, 220,390 };//设置分栏的宽度
                    float[] left = { 200, 370, 540 };//设置分栏的宽度
                    while ((status & ColumnText.NO_MORE_TEXT) == 0)
                    {
                        ct.SetSimpleColumn(right[column], 20, left[column], 795, 20, Element.ALIGN_JUSTIFIED);
                        ////设置分栏的宽度，给与参数设置，倒数第2个为每行的高度,倒数第3个为分栏的开始高度（从下往上数）
                        //倒数第4个为分栏的结束高度（从下往上数）
                        status = ct.Go();
                        if ((status & ColumnText.NO_MORE_COLUMN) != 0)//如果栏数不够，则新开一页
                        {
                            column++;
                            if (column > 2)//这里申明列数
                            {
                                document.NewPage();
                                column = 0;
                            }
                        }
                    }

                }

                catch (DocumentException de)
                {
                    Console.Error.WriteLine(de.Message);
                }
                catch (IOException ioe)
                {
                    Console.Error.WriteLine(ioe.Message);
                }
                document.Close();
                try
                {
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    MessageBox.Show("需要装PDF浏览器！生成的PDF保存在程序目录Records中(OOXX)");
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //=button3.更改numericUpDown
            if (Words.Rows.Count == 0)
            {
                MessageBox.Show("先生成单词表！");
                return;
            }
            else
            {
                Document document = new Document();
                document.AddAuthor("fakenda");
                string filename = "单词表  " + DateTime.Now.ToShortDateString() + ".pdf";
                string path = Application.StartupPath + @"\Records\" + filename;
                try
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(path, FileMode.Create));
                    total = (int)numericUpDown4.Value - (int)numericUpDown3.Value + 1;
                    //document.SetMargins(1f, 1f, 2f, 2f);
                    //document.SetMarginMirroring(true);
                    //document.GetBottom(1);
                    HeaderFooter header = new HeaderFooter(new Phrase("OOXX     " + DateTime.Now.ToShortDateString() + "         Total    " + total.ToString() + "  words"), false);
                    document.Header = header;//Header.Footer 需要在document.Open()前设置
                    //document.Footer = new HeaderFooter(new Phrase("This is page:   " ),true );

                    document.SetPageSize(iTextSharp.text.PageSize.A4);
                    document.Open();


                    for (int i = (int)numericUpDown3.Value - 1; i < (int)numericUpDown4.Value; i++)
                    {
                        Table table = DoTable(Words.Rows[i]["平假名"].ToString(), Words.Rows[i]["汉字"].ToString(), Words.Rows[i]["释义"].ToString(), Words.Rows[i]["例句"].ToString(), Words.Rows[i]["词性"].ToString(), Words.Rows[i]["单词级别"].ToString());
                        if (table == null)
                        {
                            Close();
                            return;

                        }
                        //document.Add(table);

                        if (!writer.FitsPage(table))
                        {
                            //table.DeleteLastRow();
                            //i--;
                            //table.Offset = 0;
                            document.NewPage();
                            document.Add(table);

                            //table = getTable();
                        }
                        else
                        {
                            //table.Offset = 32;
                            document.Add(table);

                        }
                    }
                }

                catch (DocumentException de)
                {
                    Console.Error.WriteLine(de.Message);
                }
                catch (IOException ioe)
                {
                    Console.Error.WriteLine(ioe.Message);
                }
                document.Close();
                try
                {
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    MessageBox.Show("需要装PDF浏览器！生成的PDF保存在程序目录Records中(OOXX)");
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //=button9.更改numericUpDown
            if (Words.Rows.Count == 0)
            {
                MessageBox.Show("先生成单词表！");
                return;
            }
            else
            {
                Document document = new Document();
                document.AddAuthor("fakenda");
                string filename = "测验卷  " + DateTime.Now.ToShortDateString() + ".pdf";
                string path = Application.StartupPath + @"\Records\" + filename;
                try
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(path, FileMode.Create));
                    total = (int)numericUpDown4.Value - (int)numericUpDown3.Value + 1;
                    //document.SetMargins(1f, 1f, 2f, 2f);
                    //document.SetMarginMirroring(true);
                    //HeaderFooter header = new HeaderFooter(new Phrase("OOXX     " + DateTime.Now.ToShortDateString() + "         Total    " + total.ToString() + "  words"), false);
                    //document.Header = header;//Header.Footer 需要在document.Open()前设置
                    //document.Footer = header;
                    document.SetPageSize(iTextSharp.text.PageSize.A4);
                    document.Open();

                    Table datatable = new Table(5);
                    datatable.CellsFitPage = true;
                    datatable.Width = 100f;//设置表整体的宽度
                    datatable.Padding = 3;
                    datatable.Spacing = 3;
                    //datatable.setBorder(Rectangle.NO_BORDER);
                    float[] headerwidths = { 25, 25, 1, 25, 25 };//设置表中列的宽度
                    datatable.Widths = headerwidths;
                    //datatable.WidthPercentage = 100;

                    BaseFont YHei = null;
                    try
                    {
                        YHei = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\msyh.TTF", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        //加载系统中字体；
                    }
                    catch
                    {
                        MessageBox.Show("抱歉，需要装微软雅黑字体先!");
                        return;
                    }
                    iTextSharp.text.Font JM = new iTextSharp.text.Font(YHei, 14, 2, iTextSharp.text.Color.BLACK);
                    iTextSharp.text.Font TT = new iTextSharp.text.Font(YHei, 16, 0, iTextSharp.text.Color.GRAY);

                    Cell cell = new Cell(new Phrase("fakenda@hotmail.com", TT));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.Leading = 30;
                    cell.Colspan = 5;
                    cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                    cell.BackgroundColor = new iTextSharp.text.Color(0xC0, 0xC0, 0xC0);
                    datatable.AddCell(cell);
                    datatable.EndHeaders();

                    datatable.DefaultCellBorderWidth = 2;
                    datatable.DefaultHorizontalAlignment = 1;
                    datatable.DefaultRowspan = 1;
                    datatable.DefaultColspan = 1;
                    datatable.DefaultHorizontalAlignment = Element.ALIGN_LEFT;

                    string key = "";

                    if (checkBox4.Checked == true)
                    {
                        key = "汉字";
                    }
                    else
                    {
                        key = "平假名";
                    }

                    for (int i = (int)numericUpDown3.Value - 1; i < (int)numericUpDown4.Value; i = i + 2)
                    {
                        cell = new Cell(new Paragraph("(" + (i + 1).ToString() + ") " + Words.Rows[i][key].ToString().Split('【')[0].Trim(), JM));//这样可以设置字体形式

                        datatable.AddCell(cell);

                        cell = new Cell(new Paragraph("", JM));
                        datatable.AddCell(cell);

                        cell = new Cell(new Paragraph("", JM));
                        cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                        cell.BackgroundColor = iTextSharp.text.Color.BLACK;
                        datatable.AddCell(cell);

                        try
                        {
                            cell = new Cell(new Paragraph("(" + (i + 2).ToString() + ") " + Words.Rows[i + 1][key].ToString().Split('【')[0].Trim(), JM));
                            datatable.AddCell(cell);

                            cell = new Cell(new Paragraph("", JM));
                            datatable.AddCell(cell);
                        }
                        catch
                        {
                            break;
                        }

                    }
                    document.Add(datatable);

                    document.NewPage();
                    document.Add(new Phrase("ANSWER:", JM));

                    PdfContentByte cb = writer.DirectContent;
                    ColumnText ct = new ColumnText(cb);

                    if (key == "平假名")
                    {
                        key = "汉字";
                    }
                    else
                    {
                        key = "平假名";
                    }
                    for (int i = (int)numericUpDown3.Value - 1; i < (int)numericUpDown4.Value; i++)
                    {
                        ct.AddText(new Paragraph("(" + (i + 1).ToString() + ") " + Words.Rows[i][key].ToString().Split('【')[0].Trim() + '\n', JM));
                    }
                    ct.Indent = 10;//缩进
                    int status = 0;
                    int column = 0;
                    float[] right = { 50, 220, 390 };//设置分栏的宽度
                    float[] left = { 200, 370, 540 };//设置分栏的宽度
                    while ((status & ColumnText.NO_MORE_TEXT) == 0)
                    {
                        ct.SetSimpleColumn(right[column], 20, left[column], 795, 20, Element.ALIGN_JUSTIFIED);
                        ////设置分栏的宽度，给与参数设置，倒数第2个为每行的高度,倒数第3个为分栏的开始高度（从下往上数）
                        //倒数第4个为分栏的结束高度（从下往上数）
                        status = ct.Go();
                        if ((status & ColumnText.NO_MORE_COLUMN) != 0)//如果栏数不够，则新开一页
                        {
                            column++;
                            if (column > 2)//这里申明列数
                            {
                                document.NewPage();
                                column = 0;
                            }
                        }
                    }

                }

                catch (DocumentException de)
                {
                    Console.Error.WriteLine(de.Message);
                }
                catch (IOException ioe)
                {
                    Console.Error.WriteLine(ioe.Message);
                }
                document.Close();
                try
                {
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    MessageBox.Show("需要装PDF浏览器！生成的PDF保存在程序目录Records中(OOXX)");
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Words.Rows.Clear();
            string sqlstr = "select 序号,平假名,汉字,释义,例句,词性,单词级别,总背诵次数,背诵重要级别,背诵标志 from "+dancibiao +" where 词性 ='";
            string sqlwt = comboBox1.Text;
            string sqlwl = comboBox2.Text;
            string sqlString1 = sqlstr + sqlwt + "'and 单词级别='" + sqlwl+"'";
            string sqlOrder = "";
            string sqlString = "";
            try
            {
                                
                //根据不同的排序选项进行选择不同的语句
                if (checkBox_sj2.Checked)
                {
                    sqlOrder = " ORDER BY right(cstr(rnd(-int(rnd(-timer())*100+序号)))*1000*Now(),2) ";
                    //随机排序函数。好东西！
                }

                else if (checkBox_nd2.Checked)
                {
                    sqlOrder = "  ORDER BY 背诵重要级别 desc";
                }

                else
                {

                }
                //

                sqlString = sqlString1 + sqlOrder + ";";
                Console.WriteLine(sqlString);
            }
            catch
            {
                MessageBox.Show("??Error！");
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
            total = Words.Rows.Count;
            numericUpDown4.Value = total;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Words.Rows.Clear();
            string sqlstr = "";
            string sqlOrder = "";
            string sqlString = "";
            if (checkBox_bj2.Checked)
            {
                sqlstr = "select 序号,平假名,汉字,释义,例句,词性,单词级别,总背诵次数,背诵重要级别,背诵标志 from "+dancibiao +" where 背诵标志=true";
            }
            else
            {
                if (DialogResult .OK  == MessageBox.Show("选择生成全部单词,确定吗?", "??", MessageBoxButtons.OKCancel))
                {
                    sqlstr = "select 序号,平假名,汉字,释义,例句,词性,单词级别,总背诵次数,背诵重要级别,背诵标志 from " + dancibiao + " ";
                }
                else
                {
                    return;
                }
            }
            try
            {

                //根据不同的排序选项进行选择不同的语句
                if (checkBox_sj2.Checked)
                {
                    sqlOrder = " ORDER BY right(cstr(rnd(-int(rnd(-timer())*100+序号)))*1000*Now(),2) ";
                    //随机排序函数。好东西！
                }

                else if (checkBox_nd2.Checked)
                {
                    sqlOrder = "  ORDER BY 背诵重要级别 desc";
                }

                else
                {

                }
                //

                sqlString = sqlstr + sqlOrder + ";";
                Console.WriteLine(sqlString);
            }
            catch
            {
                MessageBox.Show("??Error！");
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
            total = Words.Rows.Count;
            numericUpDown4.Value = total;
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Close();
        }













        
    }
}