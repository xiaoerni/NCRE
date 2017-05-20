using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLL;
using Model;
using ppt = Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Microsoft.Win32;
using SHDocVw;
using NCRE学生考试端V1._0;

/////////添加超链接
namespace NCRE学生考试端V1._0
{

    #region 超链接H卷2015年12月11日17:28:39

    public class PptHyperlinkA
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptHyperlinkA()
        {

            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();


        public void ppthyperlink(PptQuestionEntity pptinfo)
        {

            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            pptinfo.QuestionFlag = "超链接";
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[2].Hyperlinks.Count > 0)
            {
                studentrecord.ExamAnswer = pp1.Slides[2].Hyperlinks[1].Address.ToString();

                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
                {

                    studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
                }
                else
                {
                    studentrecord.Fration = "0";
                }
            }
            else
            {
                studentrecord.ExamAnswer = "未设置超链接";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion
   

    #region 超链接D卷2015年12月11日16:16:28

    public class PptHyperlinkD
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptHyperlinkD()
        {

            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTD.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();


        public void ppthyperlink(PptQuestionEntity pptinfo)
        {

            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            pptinfo.QuestionFlag = "超链接";
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[3].Hyperlinks.Count > 0)
            {
                studentrecord.ExamAnswer = pp1.Slides[3].Hyperlinks[1].Address.ToString();

                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
                {

                    studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
                }
                else
                {
                    studentrecord.Fration = "0";
                }
            }
            else
            {
                studentrecord.ExamAnswer = "未设置超链接";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region 超链接C卷，链接到第一张幻灯片上

    public class PptHyperlinkC
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptHyperlinkC()
        {

            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();


        public void ppthyperlink(PptQuestionEntity pptinfo)
        {

            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            //pptinfo.PaperType = dt1.Rows[0]["PaperType"].ToString();

            pptinfo.QuestionFlag = "超链接";
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[5].Hyperlinks.Count > 0)
            {
                studentrecord.ExamAnswer = pp1.Slides[5].Hyperlinks[1].SubAddress.ToString();

                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
                {

                    studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
                }
                else
                {
                    studentrecord.Fration = "0";
                }
            }
            else
            {
                studentrecord.ExamAnswer = "未设置超链接";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region 超链接G卷，链接到第一张幻灯片上2015年12月11日16:58:34

    public class PptHyperlinkG
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptHyperlinkG()
        {

            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTG.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();


        public void ppthyperlink(PptQuestionEntity pptinfo)
        {

            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            pptinfo.QuestionFlag = "超链接";
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            if (pp1.Slides[4].Hyperlinks.Count>0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                pptquestionbll.ReturnScore(studentrecord);

            }
            else
            {
                //将考生答案传递给studentrecord实体
                if (pp1.Slides[4].Hyperlinks.Count > 0)
                {
                    studentrecord.ExamAnswer = pp1.Slides[4].Hyperlinks[1].SubAddress.Substring(0, 6);

                    if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
                    {

                        studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
                    }
                    else
                    {
                        studentrecord.Fration = "0";
                    }
                }
                else
                {
                    studentrecord.ExamAnswer = "未设置超链接";
                    studentrecord.Fration = "0";
                }
                pptquestionbll.ReturnScore(studentrecord);
            }
        }
    }
    #endregion

    #region 超链接H卷2015年12月11日17:12:24

    public class PptHyperlinkH
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptHyperlinkH()
        {

            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTH.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();


        public void ppthyperlink(PptQuestionEntity pptinfo)
        {

            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            pptinfo.QuestionFlag = "超链接";
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[3].Hyperlinks.Count > 0)
            {
                studentrecord.ExamAnswer = pp1.Slides[3].Hyperlinks[1].Address.ToString();

                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
                {

                    studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
                }
                else
                {
                    studentrecord.Fration = "0";
                }
            }
            else
            {
                studentrecord.ExamAnswer = "未设置超链接";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
}
    #endregion



