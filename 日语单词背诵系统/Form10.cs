using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DotNetSpeech;
using System.Threading;
using System.Data.OleDb;

namespace 日语单词背诵系统
{
    public partial class Form10 : Form
    {
        string voiced = form1.voice;
        string dancibiao = form1.G_danci;
        string jihuabiao = form1.G_jihua;

        int volume = 100;
        int rate = -2;


        public Form10()
        {
            InitializeComponent();
        }

        public void Read(string readtxt)
        {
            SpVoiceClass voice = new SpVoiceClass();

            try
            {

                voice.Voice = voice.GetVoices("name=" + voiced, string.Empty).Item(0);

                voice.Rate = rate ;
            }
            catch
            {
                MessageBox.Show("语音插件没有安装或正确配置，不能使用！");
                return;
            }

            try
            {
                voice.Speak(readtxt, SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak);
            }
            catch
            {
                MessageBox.Show("未注册的语音包!请重新安装!");
            }

            voice = null;
            GC.Collect();

        }

        public void enduce(string endutxt)
        {
            SpVoiceClass speech = new SpVoiceClass();
            try
            {
                speech.Voice = speech.GetVoices("name=" + voiced, string.Empty).Item(0);
                speech.Rate = rate ;
            }
            catch
            {
                MessageBox.Show("语音插件没有安装或正确配置，不能使用！");
                return;
            }
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "所有文件 (*.*)|*.*|WAV 格式文件 (*.wav)|*.wav";
                sfd.Title = "保存到 wave 文件";
                sfd.FilterIndex = 2;
                sfd.FileName = endutxt;
                sfd.RestoreDirectory = true;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    SpeechStreamFileMode SpFileMode = SpeechStreamFileMode.SSFMCreateForWrite;
                    SpFileStream SpFileStream = new SpFileStream();
                    SpFileStream.Open(sfd.FileName, SpFileMode, false);
                    speech.AudioOutputStream = SpFileStream;
                    speech.Rate =rate ;
                    speech.Volume = volume ;
                    speech.Speak(endutxt, SpeechVoiceSpeakFlags.SVSFlagsAsync);
                    speech.WaitUntilDone(Timeout.Infinite);
                    SpFileStream.Close();
                }
            }
            catch
            {
                MessageBox.Show("导出Wav文件出错！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            speech = null;
            GC.Collect();
        }



        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (textBoxX1.Text == "")
            {
                return;
            }
            else
            {
                Read(textBoxX1.Text);
            }

        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (textBoxX1.Text == "")
            {
                return;
            }
            else
            {
                enduce (textBoxX1.Text);
            }
            textBoxX1.Text = "";
            textBoxX1.Focus();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            Conn conn = new Conn();
            OleDbConnection connection = conn.CreatConn();
            connection.Open();
            string key = textBoxX1.Text;
            if (key == "")
            {
                return;
            }
            else
            {
                //key = key.Split('【')[0].Trim();
            }

            string sqlstring = "select 释义 from " + jihuabiao + " where 平假名= '" + key.ToString() + "' or 汉字='" + key.ToString() + "'"; ;
            try
            {
                OleDbCommand myCommand = new OleDbCommand(sqlstring, connection);
                OleDbDataReader myReader = myCommand.ExecuteReader();
                while (myReader.Read())
                {

                }
                myReader.Close();
                connection.Close();
            }
            catch
            {
                MessageBox.Show("发生错误", "提示！", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);

            }
        }

        private void Form10_Enter(object sender, EventArgs e)
        {
            textBoxX1.Focus();
        }
    }
}
