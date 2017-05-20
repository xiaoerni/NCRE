﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Model;
using System.Data.SqlClient;
using BLL;
using System.Threading;
using Microsoft.Win32;
using System.Diagnostics;
using NCRE学生考试端V1._0.选择题;
using NCRE学生考试端V1._0.悬浮框;
using System.Runtime.InteropServices;


namespace NCRE学生考试端V1._0
{
    public partial class frmMain2 : Form
    {
        public static Form frmmain2;
        public frmMain2()
        {
            InitializeComponent();
            frmmain2 = this;

        }

        //实例化一个学生实体
        public StudentInfoEntity studentinfo = new StudentInfoEntity();

        /// <summary>
        /// 加载学生信息
        /// </summary>
        public static StudentInfoEntity st = new StudentInfoEntity();
        private StudentInfoEntity InitStudent(string studentId, string collegeId)
        {
            st.studentID = studentId;
            st.CollegeID = collegeId;
            return st;
        }


        private void frmMain2_Load(object sender, EventArgs e)
        {
            //确定操作题窗体的位置--周洲--2015-12-13
            this.Top = 0;
            this.Left = 0;
            Rectangle ScreenArea = System.Windows.Forms.Screen.GetWorkingArea(this);
            this.Height = ScreenArea.Height;

            //确定所有的显示题txt的高度
            txtWin.Height = ScreenArea.Height - 200;
            txtWord.Height = ScreenArea.Height - 200;
            txtPPT.Height = ScreenArea.Height - 200;
            txtIE.Height = ScreenArea.Height - 200;
            txtExcel.Height = ScreenArea.Height - 200;
            //上一步不可用
            button3.Visible = false;

            #region 利用全局变量，从题库中加载word试题--周洲--2015年11月21日
            //定义一个word助手类
            WordLoadinfo wordhelper = new WordLoadinfo();
            //定义一个题库类传递Papertype
            WordQuestionEntity wordinfo = new WordQuestionEntity();
            wordinfo.PaperType = MyInfo.MyPaperType();
            //调用word试题load的方法
            DataTable worddt = wordhelper.LoadQuestionContent(wordinfo);
            //将从数据库中取出的字段赋给一个字符串
            string newLine = null;
            //循环DataTable取出里面的值
            for (int i = 0; i < worddt.Rows.Count; i++)
            {
                newLine += worddt.Rows[i]["QuestionContent"].ToString();

            }
            //让字符串按照规律 赋给文本框 
            string[] s = newLine.Split('。');
            for (int i = 0; i < s.Length; i++)
            {
                txtWord.Text += s[i] + "\r\n";
            }
            #endregion

            #region 利用全局变量，从题库中加载Excel试题--周洲--2015年11月21日

            ExcelLoading excelhelper = new ExcelLoading();
            ExcelEntityBLL excelentitybll = new ExcelEntityBLL();
            ExcelQuestionEntity excelinfo = new ExcelQuestionEntity();

            excelinfo.PaperType = MyInfo.MyPaperType();
            DataTable exceldt = excelhelper.LoadQuestionContent(excelinfo);
            //将从数据库中取出的字段赋给一个字符串
            string newExcelContent = null;
            //循环DataTable取出里面的值
            for (int i = 0; i < exceldt.Rows.Count; i++)
            {
                newExcelContent += exceldt.Rows[i]["QuestionContent"].ToString();
            }
            //让字符串按照规律 赋给文本框 
            string[] sE = newExcelContent.Split('。');
            for (int i = 0; i < sE.Length; i++)
            {
                txtExcel.Text += sE[i] + "\r\n";
            }
            #endregion

            #region 利用全局变量，从题库中加载windows试题--周洲--2015年11月21日
            //定义一个windows助手类
            WindowsLoadInfo windowsHelper = new WindowsLoadInfo();
            //定义一个Windows题库类
            WinQuestionEntity wininfo = new WinQuestionEntity();
            wininfo.paperType = MyInfo.MyPaperType();
            //调用WindowsLoadInfo中的LoadQuestionContent方法
            DataTable winQuestionDt = new DataTable();
            winQuestionDt = windowsHelper.LoadQuestionContent(wininfo);
            //将从数据库中取出的字段赋给一个字符串
            string newWinContent = null;
            //循环winQuestionDt取出里面的所有值
            for (int i = 0; i < winQuestionDt.Rows.Count; i++)
            {

                newWinContent += winQuestionDt.Rows[i]["questionContent"].ToString();
            }
            //让字符串按照规律 赋给文本框 
            string[] sW = newWinContent.Split('。');
            for (int i = 0; i < sW.Length; i++)
            {
                txtWin.Text += sW[i] + "\r\n";
            }
            #endregion

            #region 利用全局变量，从题库中加载IE试题--周洲--2015年11月21日
            //定义一个windows助手类
            IELoadInfo ieHelper = new IELoadInfo();
            //定义一个Windows题库类
            IEQuestionEntity ieinfo = new IEQuestionEntity();
            //调用WindowsLoadInfo中的LoadQuestionContent方法
            DataTable ieQuestionDt = new DataTable();
            ieinfo.paperType = MyInfo.MyPaperType();
            ieQuestionDt = ieHelper.LoadQuestionContent(ieinfo);
            //将从数据库中取出的字段赋给一个字符串
            string newIEContent = null;
            //循环ieQuestionDt取出里面的所有值
            for (int i = 0; i < ieQuestionDt.Rows.Count; i++)
            {
                newIEContent += ieQuestionDt.Rows[i]["questionContent"].ToString();
            }
            //让字符串按照规律 赋给文本框 
            string[] sIE = newIEContent.Split('。');
            for (int i = 0; i < sIE.Length; i++)
            {
                txtIE.Text += sIE[i] + "\r\n";
            }
            #endregion

            #region 利用全局变量，从题库中加载PPT试题--周洲--2015年11月21日
            //定义一个PPT助手类
            PptLoadinfo ppthelper = new PptLoadinfo();
            //定义一个PPT题库类
            PptQuestionEntity pptinfo = new PptQuestionEntity();
            pptinfo.PaperType = MyInfo.MyPaperType();
            //调用PPTLoadInfo中的LoadQuestionContent方法
            DataTable pptDt = new DataTable();
            pptDt = ppthelper.LoadQuestionContent(pptinfo);
            //将从数据库中取出的字段赋给一个字符串
            string newPPTContent = null;
            //循环ieQuestionDt取出里面的所有值
            for (int i = 0; i < pptDt.Rows.Count; i++)
            {
                newPPTContent += pptDt.Rows[i]["QuestionContent"].ToString();
            }
            //让字符串按照规律 赋给文本框 
            string[] sPPT = newPPTContent.Split('。');
            for (int i = 0; i < sPPT.Length; i++)
            {
                txtPPT.Text += sPPT[i] + "\r\n";
            }
            #endregion
        }

        /// <summary>
        /// 上一题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            //判断当前页面是哪个，切换到上个题
            if (txtWord.Visible == true)
            {
                button2.Visible=true;
                button3.Visible = false;
                button1.Visible = true;
                //  当前显示操作系统题

                label1.ForeColor = System.Drawing.Color.Red;
                label2.ForeColor = System.Drawing.Color.Gray;
                label3.ForeColor = System.Drawing.Color.Gray;
                label4.ForeColor = System.Drawing.Color.Gray;
                label5.ForeColor = System.Drawing.Color.Gray;
                txtWin.Visible = true;
                txtWord.Visible = false;
                txtExcel.Visible = false;
                txtIE.Visible = false;
                txtPPT.Visible = false;
            }

            else if (txtExcel.Visible == true)
            {
                button1.Visible = false;
                button2.Visible = true;
                button3.Visible = true;
                //当前显示word题目
                label1.ForeColor = System.Drawing.Color.Gray;
                label2.ForeColor = System.Drawing.Color.Red;
                label3.ForeColor = System.Drawing.Color.Gray;
                label4.ForeColor = System.Drawing.Color.Gray;
                label5.ForeColor = System.Drawing.Color.Gray;
                txtWord.Visible = true;
                txtExcel.Visible = false;
                txtPPT.Visible = false;
                txtIE.Visible = false;
                txtWin.Visible = false;
            }
            else if (txtPPT.Visible == true)
            {
                button1.Visible = false;
                button2.Visible = true;
                button3.Visible = true;
                //当前显示Excel题
                label1.ForeColor = System.Drawing.Color.Gray;
                label2.ForeColor = System.Drawing.Color.Gray;
                label3.ForeColor = System.Drawing.Color.Red;
                label4.ForeColor = System.Drawing.Color.Gray;
                label5.ForeColor = System.Drawing.Color.Gray;
                txtExcel.Visible = true;
                txtPPT.Visible = false;
                txtIE.Visible = false;
                txtWord.Visible = false;
                txtWin.Visible = false;
            }
            else if (txtIE.Visible == true)
            {
                button1.Visible = false;
                button2.Visible = true;
                button3.Visible = true;
                //当前显示ppt的题
                label1.ForeColor = System.Drawing.Color.Gray;
                label2.ForeColor = System.Drawing.Color.Gray;
                label3.ForeColor = System.Drawing.Color.Gray;
                label4.ForeColor = System.Drawing.Color.Red;
                label5.ForeColor = System.Drawing.Color.Gray;
                txtPPT.Visible = true;
                txtIE.Visible = false;
                txtExcel.Visible = false;
                txtWord.Visible = false;
                txtWin.Visible = false;
            }
        }


        /// <summary>
        /// 下一题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            //判断当前页面是哪个，切换到下个题
            if (txtWin.Visible == true)
            {
                button3.Visible = true;
                button2.Visible = true;
                button1.Visible = false;
                //当前显示word题
                label1.ForeColor = System.Drawing.Color.Gray;
                label2.ForeColor = System.Drawing.Color.Red;
                label3.ForeColor = System.Drawing.Color.Gray;
                label4.ForeColor = System.Drawing.Color.Gray;
                label5.ForeColor = System.Drawing.Color.Gray;
                txtWord.Visible = true;
                txtWin.Visible = false;
                txtPPT.Visible = false;
                txtIE.Visible = false;
                txtExcel.Visible = false;
            }

            else if (txtWord.Visible == true)
            {
                button1.Visible = false;
                button3.Visible = true;
                button2.Visible = true;
                //当前显示Excel题
                label1.ForeColor = System.Drawing.Color.Gray;
                label2.ForeColor = System.Drawing.Color.Gray;
                label3.ForeColor = System.Drawing.Color.Red;
                label4.ForeColor = System.Drawing.Color.Gray;
                label5.ForeColor = System.Drawing.Color.Gray;
                txtExcel.Visible = true;
                txtWord.Visible = false;
                txtWin.Visible = false;
                txtIE.Visible = false;
                txtPPT.Visible = false;
            }

            else if (txtExcel.Visible == true)
            {
                button1.Visible = false;
                button3.Visible = true;
                button2.Visible = true;
                //当前显示PPT题
                label1.ForeColor = System.Drawing.Color.Gray;
                label2.ForeColor = System.Drawing.Color.Gray;
                label3.ForeColor = System.Drawing.Color.Gray;
                label4.ForeColor = System.Drawing.Color.Red;
                label5.ForeColor = System.Drawing.Color.Gray;
                txtPPT.Visible = true;
                txtExcel.Visible = false;
                txtWord.Visible = false;
                txtIE.Visible = false;
                txtWin.Visible = false;
            }
            else if (txtPPT.Visible == true)
            {
                button1.Visible = false;
                button2.Visible = false;
                button3.Visible = true;
                //当前显示IE题
                //打开本地网页-韩梦甜-2014-12-6
                Process.Start(@"D:\计算机一级考生文件\netkt\百度一下，你就知道.html"); 
                label1.ForeColor = System.Drawing.Color.Gray;
                label2.ForeColor = System.Drawing.Color.Gray;
                label3.ForeColor = System.Drawing.Color.Gray;
                label4.ForeColor = System.Drawing.Color.Gray;
                label5.ForeColor = System.Drawing.Color.Red;

                txtIE.Visible = true;
                txtPPT.Visible = false;
                txtPPT.Visible = false;
                txtWord.Visible = false;
                txtExcel.Visible = false;

            }
        }

        private void picBig_Click(object sender, EventArgs e)
        {
            //放大字体
            if (txtWin.Font.Size <= 22)
            {

                int fontWinSize = int.Parse(this.txtWin.Font.Size.ToString()) + 2;
                txtWin.Font=new Font (txtWin.Font.Name , fontWinSize);
                txtExcel.Font = new Font(txtExcel.Font.Name, fontWinSize);
                txtIE.Font = new Font(txtWin.Font.Name, fontWinSize);
                txtPPT.Font = new Font(txtWin.Font.Name, fontWinSize);
                txtWord.Font = new Font(txtWin.Font.Name, fontWinSize);
                picSmall.Visible = true;
            }
            else
            {
                picBig.Visible = false;
                picSmall.Visible = true;
            }
          
        }

        private void picSmall_Click(object sender, EventArgs e)
        {
            //缩小字体
            if (txtWin.Font.Size >12)
            {
                int fontWinSize = int.Parse(this.txtWin.Font.Size.ToString()) - 2;
                txtWin.Font = new Font(txtWin.Font.Name, fontWinSize);
                txtExcel.Font = new Font(txtExcel.Font.Name, fontWinSize);
                txtIE.Font = new Font(txtWin.Font.Name, fontWinSize);
                txtPPT.Font = new Font(txtWin.Font.Name, fontWinSize);
                txtWord.Font = new Font(txtWin.Font.Name, fontWinSize);
                picBig.Visible = true;
            }
            else
            {
                picSmall.Visible = false;
                picBig.Visible = true;
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            //调回到选择题
            this.Hide();
            FrmMain frmmain = new FrmMain();
            frmmain.ShowDialog();
        }

        private void frmMain2_FormOnclosing(object sender, FormClosedEventArgs e)
        {
          
            
        }
      

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            e.Cancel = true;
            frmxuanfukuang frmxfk = new frmxuanfukuang();
            frmxfk.showbtn();
            this.Hide();
        }
        private void label5_Click(object sender, EventArgs e)
        {
            //当前显示IE题
            label1.ForeColor = System.Drawing.Color.Black;
            label2.ForeColor = System.Drawing.Color.Black;
            label3.ForeColor = System.Drawing.Color.Black;
            label4.ForeColor = System.Drawing.Color.Black;
            label5.ForeColor = System.Drawing.Color.Red;

            txtWord.Visible = false;
            txtWin.Visible = false ;
            txtPPT.Visible = false;
            txtIE.Visible = true  ;
            txtExcel.Visible = false;
        }



        private void frmMain2_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {
            //当前显示win题
            label1.ForeColor = System.Drawing.Color.Red;
            label2.ForeColor = System.Drawing.Color.Black;
            label3.ForeColor = System.Drawing.Color.Black;
            label4.ForeColor = System.Drawing.Color.Black;
            label5.ForeColor = System.Drawing.Color.Black;

            txtWord.Visible = false ;
            txtWin.Visible = true;
            txtPPT.Visible = false;
            txtIE.Visible = false;
            txtExcel.Visible = false;
        }

        private void label2_Click(object sender, EventArgs e)
        {
            //当前显示word题
            label1.ForeColor = System.Drawing.Color.Black ;
            label2.ForeColor = System.Drawing.Color.Red;
            label3.ForeColor = System.Drawing.Color.Black;
            label4.ForeColor = System.Drawing.Color.Black;
            label5.ForeColor = System.Drawing.Color.Black;
            
            txtWord.Visible = true;
            txtWin.Visible = false;
            txtPPT.Visible = false;
            txtIE.Visible = false;
            txtExcel.Visible = false;
        }

        private void label3_Click(object sender, EventArgs e)
        {
            //当前显示excel题
            label1.ForeColor = System.Drawing.Color.Black;
            label2.ForeColor = System.Drawing.Color.Black;
            label3.ForeColor = System.Drawing.Color.Red;
            label4.ForeColor = System.Drawing.Color.Black;
            label5.ForeColor = System.Drawing.Color.Black;

            txtWord.Visible = false;
            txtWin.Visible = false;
            txtPPT.Visible = false;
            txtIE.Visible = false;
            txtExcel.Visible = true;
        }

        private void label4_Click(object sender, EventArgs e)
        {
            //当前显示ppt题
            label1.ForeColor = System.Drawing.Color.Black;
            label2.ForeColor = System.Drawing.Color.Black;
            label3.ForeColor = System.Drawing.Color.Black;
            label4.ForeColor = System.Drawing.Color.Red;
            label5.ForeColor = System.Drawing.Color.Black;

            txtWord.Visible = false;
            txtWin.Visible = false;
            txtPPT.Visible = true ;
            txtIE.Visible = false;
            txtExcel.Visible = false ;
        }

        
    }
}
