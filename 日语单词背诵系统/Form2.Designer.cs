namespace 日语单词背诵系统
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.trackBar1 = new System.Windows.Forms.TrackBar();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.textBox6 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.label2.Location = new System.Drawing.Point(28, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(42, 21);
            this.label2.TabIndex = 11;
            this.label2.Text = "假名";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.label3.Location = new System.Drawing.Point(28, 70);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(42, 21);
            this.label3.TabIndex = 12;
            this.label3.Text = "汉字";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.label4.Location = new System.Drawing.Point(28, 116);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(42, 21);
            this.label4.TabIndex = 13;
            this.label4.Text = "释义";
            // 
            // textBox1
            // 
            this.errorProvider1.SetIconPadding(this.textBox1, 20);
            this.textBox1.ImeMode = System.Windows.Forms.ImeMode.On;
            this.textBox1.Location = new System.Drawing.Point(101, 21);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(145, 29);
            this.textBox1.TabIndex = 1;
            this.toolTip1.SetToolTip(this.textBox1, "不能为空");
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(101, 68);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 29);
            this.textBox2.TabIndex = 2;
            // 
            // textBox3
            // 
            this.textBox3.AcceptsReturn = true;
            this.textBox3.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox3.Location = new System.Drawing.Point(101, 114);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox3.Size = new System.Drawing.Size(292, 53);
            this.textBox3.TabIndex = 3;
            this.toolTip1.SetToolTip(this.textBox3, "不能为空");
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.label5.Location = new System.Drawing.Point(28, 236);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(42, 21);
            this.label5.TabIndex = 14;
            this.label5.Text = "例句";
            // 
            // textBox4
            // 
            this.textBox4.AcceptsReturn = true;
            this.textBox4.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox4.Location = new System.Drawing.Point(101, 238);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox4.Size = new System.Drawing.Size(292, 58);
            this.textBox4.TabIndex = 6;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Maroon;
            this.button1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(465, 108);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 32);
            this.button1.TabIndex = 7;
            this.button1.Text = "保存";
            this.toolTip1.SetToolTip(this.button1, "保存单条记录");
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // trackBar1
            // 
            this.trackBar1.Location = new System.Drawing.Point(447, 31);
            this.trackBar1.Name = "trackBar1";
            this.trackBar1.Size = new System.Drawing.Size(104, 45);
            this.trackBar1.TabIndex = 10;
            this.trackBar1.Scroll += new System.EventHandler(this.trackBar1_Scroll);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Maroon;
            this.button2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button2.Location = new System.Drawing.Point(465, 167);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 32);
            this.button2.TabIndex = 8;
            this.button2.Text = "清空";
            this.toolTip1.SetToolTip(this.button2, "当前记录重新写");
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.Maroon;
            this.button3.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button3.Location = new System.Drawing.Point(465, 226);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 32);
            this.button3.TabIndex = 9;
            this.button3.Text = "上传";
            this.toolTip1.SetToolTip(this.button3, "把所有记录上传");
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "名词",
            "Ⅰ类动词",
            "Ⅱ类动词",
            "い形",
            "な形",
            "副词",
            "接续词"});
            this.comboBox1.Location = new System.Drawing.Point(101, 188);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(67, 29);
            this.comboBox1.TabIndex = 4;
            this.comboBox1.Text = "名词";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.label6.Location = new System.Drawing.Point(28, 191);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 21);
            this.label6.TabIndex = 15;
            this.label6.Text = "词性";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.label7.Location = new System.Drawing.Point(268, 191);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(42, 21);
            this.label7.TabIndex = 17;
            this.label7.Text = "等级";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox2.Location = new System.Drawing.Point(317, 188);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(43, 29);
            this.comboBox2.TabIndex = 5;
            this.comboBox2.Text = "3";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.comboBox2);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.textBox3);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.textBox4);
            this.groupBox1.Cursor = System.Windows.Forms.Cursors.Default;
            this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.groupBox1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(418, 313);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "1/50";
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(272, 21);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(121, 29);
            this.textBox5.TabIndex = 18;
            this.toolTip1.SetToolTip(this.textBox5, "这里写音标！");
            // 
            // errorProvider1
            // 
            this.errorProvider1.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.AlwaysBlink;
            this.errorProvider1.ContainerControl = this;
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(458, 277);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(93, 21);
            this.textBox6.TabIndex = 19;
            this.textBox6.Text = "①②③④⑤⑥⑦";
            this.toolTip1.SetToolTip(this.textBox6, "可以从这里复制序号！");
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(580, 337);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.trackBar1);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form2";

            this.Activated += new System.EventHandler(this.Form2_Activated);
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TrackBar trackBar1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        internal System.Windows.Forms.Label label2;
        internal System.Windows.Forms.Label label3;
        internal System.Windows.Forms.Label label4;
        internal System.Windows.Forms.TextBox textBox1;
        internal System.Windows.Forms.TextBox textBox2;
        internal System.Windows.Forms.TextBox textBox3;
        internal System.Windows.Forms.ComboBox comboBox1;
        internal System.Windows.Forms.Label label6;
        internal System.Windows.Forms.Label label7;
        internal System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.GroupBox groupBox1;
        internal System.Windows.Forms.TextBox textBox5;
        internal System.Windows.Forms.TextBox textBox6;
    }
}