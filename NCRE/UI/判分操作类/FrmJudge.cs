using BLL;
using Model;
using NCRE学生考试端V1._0.PPT操作题类;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NCRE学生考试端V1._0.悬浮框;


namespace NCRE学生考试端V1._0
{
    public partial class FrmJudge : Form
    {
        public static Form frmjudge;
       
        public FrmJudge()
        { 
            InitializeComponent();
            frmjudge = this;
        }
    
        private void button1_Click(object sender, EventArgs e)
        {
            //frmPro frmpro = new frmPro();
            //frmpro.Show();

            

            ////获取学号
            string studentid = FrmLogin.studentID;
            StudentInfoEntity Studentinfo = new StudentInfoEntity();
            Studentinfo.studentID = studentid;
            TypeSumFrationBLL typesumfration = new TypeSumFrationBLL();
            DataTable dt = typesumfration.studentIDscore(Studentinfo);



            if (dt.Columns.Count > 1)
            {
                MessageBox.Show("该试卷已经判分，不可以重复判分");

                //FrmJudge judge = new FrmJudge();
                this.Close();
                //frmxuanfukuang frmxfk = new frmxuanfukuang();

            }



            else
            {

                #region Word判分  2015年11月23日
                txtMsg.Text = "\r\nword开始判分.....";

                txtMsg.Text = "\r\nword正在判分，请耐心等待.......";

                WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
                WordQuestionEntity wordinfo = new WordQuestionEntity();
                StudentInfoEntity studentinfo = new StudentInfoEntity();
                wordinfo.PaperType = MyInfo.MyPaperType();
                switch (MyInfo.MyPaperType().Trim())
                {
                    case "A":
                        //判断表格C是否存在
                        string filePath = @"D:\计算机一级考生文件\Wordkt\bgA.docx";
                        if (MyInfo.exitsDoc(filePath))
                        {
                            WordAQuestionFlag wordquestionflag = new WordAQuestionFlag();
                            wordquestionflag.SwitchQuestionFlag(wordinfo);
                        }
                        break;
                    case "B":
                        //判断表格C是否存在
                        string filePathB = @"D:\计算机一级考生文件\Wordkt\bgB.docx";
                        if (MyInfo.exitsDoc(filePathB))
                        {
                            WordBQuestionFlag wordBquestionflag = new WordBQuestionFlag();
                            wordBquestionflag.SwitchQuestionFlagB(wordinfo);
                        }
                        break;
                    case "C":
                        //判断表格C是否存在
                        string filePathC = @"D:\计算机一级考生文件\Wordkt\bgC.docx";
                        if (MyInfo.exitsDoc(filePathC))
                        {
                            WordCQuestionFlag wordCquestionflag = new WordCQuestionFlag();
                            wordCquestionflag.SwitchQuestionFlagC(wordinfo);
                        }
                        break;
                    case "D":
                        //判断表格C是否存在
                        string filePathD = @"D:\计算机一级考生文件\Wordkt\bgD.docx";
                        if (MyInfo.exitsDoc(filePathD))
                        {
                            WordDQuestionFlag wordDquestionflag = new WordDQuestionFlag();
                            wordDquestionflag.SwitchQuestionFlagD(wordinfo);
                        }
                        break;
                    case "E":
                        //判断表格C是否存在
                        string filePathE = @"D:\计算机一级考生文件\Wordkt\bgE.docx";
                        if (MyInfo.exitsDoc(filePathE))
                        {
                            WordEQuestionFlag wordEquestionflag = new WordEQuestionFlag();
                            wordEquestionflag.SwitchQuestionFlagE(wordinfo);
                        }
                        break;
                    case "F":
                        //判断表格C是否存在
                        string filePathF = @"D:\计算机一级考生文件\Wordkt\bgF.docx";
                        if (MyInfo.exitsDoc(filePathF))
                        {
                            WordFQuestionFlag wordFquestionflag = new WordFQuestionFlag();
                            wordFquestionflag.SwitchQuestionFlagF(wordinfo);
                        }
                        break;
                    case "G":
                        //判断表格C是否存在
                        string filePathG = @"D:\计算机一级考生文件\Wordkt\bgG.docx";
                        if (MyInfo.exitsDoc(filePathG))
                        {
                            WordGQuestionFlag wordGquestionflag = new WordGQuestionFlag();
                            wordGquestionflag.SwitchQuestionFlagG(wordinfo);
                        }
                        break;
                    case "H":
                        //判断表格C是否存在
                        string filePathH = @"D:\计算机一级考生文件\Wordkt\bgH.docx";
                        if (MyInfo.exitsDoc(filePathH))
                        {
                            WordHQuestionFlag wordHquestionflag = new WordHQuestionFlag();
                            wordHquestionflag.SwitchQuestionFlagH(wordinfo);
                        }
                        break;
                    default:
                        break;

                }
                //TypeSumFrationBLL typesumfration = new TypeSumFrationBLL();
                typesumfration.WordSumFration(Studentinfo);
                //MessageBox.Show("word判分成功！");
                txtMsg.Text += "\r\nword判分成功！";
                #endregion

                #region PPT判分 2015年11月29日
                txtMsg.Text += "\r\nPPT开始判分....";
                //PPT判分
                PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
                PptQuestionEntity pptquestion = new PptQuestionEntity();
                pptquestion.PaperType = MyInfo.MyPaperType();

                PptQuestionFlag pptquestionflag = new PptQuestionFlag();
                switch (MyInfo.MyPaperType().Trim())
                {
                    case "A":
                        pptquestionflag.PptSwitchQuestionFlagA(pptquestion);
                        break;
                    case "B":
                        pptquestionflag.PptSwitchQuestionFlagB(pptquestion);
                        break;
                    case "C":
                        pptquestionflag.PptSwitchQuestionFlagC(pptquestion);
                        break;
                    case "D":
                        pptquestionflag.SwitchQuestionFlagD(pptquestion);
                        break;
                    case "E":
                        pptquestionflag.SwitchQuestionFlagE(pptquestion);
                        break;
                    case "F":
                        pptquestionflag.SwitchQuestionFlagF(pptquestion);
                        break;
                    case "G":
                        pptquestionflag.SwitchQuestionFlagG(pptquestion);
                        break;
                    case "H":
                        pptquestionflag.SwitchQuestionFlagH(pptquestion);
                        break;
                    default:
                        break;
                }
                ////计算PPT题型总分-王荣晓-2015年11月30日
                typesumfration.PPTSumFration(Studentinfo);
                txtMsg.Text += "\r\nPPT判分成功！";
                #endregion

                #region IE判分 2015年11月23日
                txtMsg.Text += "\r\nIE开始判分....";
                IEQuestionEntity ieinfo = new IEQuestionEntity();
                ieinfo.paperType = MyInfo.MyPaperType();
                IEQuestionFlag iequestionflag = new IEQuestionFlag();
                iequestionflag.SwitchQuestionFlag(ieinfo);
                //计算PPT题型总分-王荣晓-2015年11月30日
                typesumfration.IESumFration(Studentinfo);
                txtMsg.Text += "\r\nIE判分成功！";

                #endregion

                #region Windows判分 2015年11月23日
                txtMsg.Text += "\r\nWindows开始判分....";
                WinQuestionEntity wininfo = new WinQuestionEntity();
                wininfo.paperType = MyInfo.MyPaperType();

                WindowsQuestionFlag windowsQuestionFlag = new WindowsQuestionFlag();
                windowsQuestionFlag.SwitchQuestionFlag(wininfo);
                //计算windows题型总分-王荣晓-2015年11月30日
                typesumfration.windowsSumFration(Studentinfo);
                txtMsg.Text += "\r\nWindows判分成功！";
                #endregion

                #region Excel判分 2015年11月23日
                txtMsg.Text += "\r\nExcel开始判分......";
                //MessageBox.Show("Excel开始判分.....");
                ExcelQuestionEntity excelinfo = new ExcelQuestionEntity();
                ExcelJudgeHelper excelhelper = new ExcelJudgeHelper();
                excelhelper.SelectJudge(excelinfo);
                // MessageBox.Show("Excel判分成功！");
                //计算Excel题型总分-王荣晓-2015年11月30日
                typesumfration.ExcelSumFration(Studentinfo);

                txtMsg.Text += "\r\nExcel判分成功！";
                //计算选择题型的总分—李芬-2015年11月30日
                typesumfration.SelectsumFration(Studentinfo);
                //计算一个学生一套试卷的总分-李芬-2015年11月30日
                typesumfration.SumFration(Studentinfo);
                txtMsg.Text += "\r\n判分成功！请选择退出....";
                #endregion

                button1.Enabled = true;
           // 显示学生的成绩
            MessageBox.Show("判分成功！请选择退出！");
             this.Close();
              }

         

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            //退出
            frmMain2.frmmain2.Show();
            this.Close();

        }

        private void FrmJudge_Load(object sender, EventArgs e)
        {
            button1.Enabled = false;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            frmShowMessage frmshowmessage = new frmShowMessage();
            frmshowmessage.Show();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //if (progressBar1.Value<progressBar1.Maximum)
            //{
            //    progressBar1.Value++;
            //}
            //else
            //{
            //    timer1.Enabled = false;
            //}
        }

    }
}
