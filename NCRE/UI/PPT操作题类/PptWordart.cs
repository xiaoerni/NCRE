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

namespace NCRE学生考试端V1._0
{
    #region A卷的艺术字2015年12月11日17:30:23
    public class PptWordartA
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptWordartA()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "艺术字样式";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            
            if (pp1.Slides.Count.ToString() == "8")
            {
                studentrecord.ExamAnswer = pp1.Slides[8].Shapes[1].TextFrame.TextRange.Font.Color.RGB.ToString();
            
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
                studentrecord.ExamAnswer = "未添加新幻灯片";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region B卷的艺术字2015年12月11日17:43:08
    public class PptWordartB
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptWordartB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void WordartType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "艺术字样式";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体


            if (pp1.Slides.Count.ToString() == "8")
            {
                studentrecord.ExamAnswer = pp1.Slides[8].Shapes[1].TextFrame.TextRange.Font.Color.RGB.ToString();
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
                studentrecord.ExamAnswer = "未添加新幻灯片";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region E卷的艺术字2015年12月11日16:36:34
    public class PptWordartE
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptWordartE()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTE.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "艺术字样式";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体

            if (pp1.Slides.Count.ToString() == "5")
            {
                studentrecord.ExamAnswer = pp1.Slides[5].Shapes[1].TextFrame.TextRange.Font.Color.RGB.ToString();
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
                studentrecord.ExamAnswer = "未添加新幻灯片";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region H卷的艺术字2015年12月11日17:05:13
    public class PptWordartH
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptWordartH()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTH.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "艺术字样式";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体

            if (pp1.Slides[1].Shapes.Count.ToString() == "2")
            {
                studentrecord.ExamAnswer = pp1.Slides[1].Shapes[2].TextFrame.TextRange.Font.Color.RGB.ToString();
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
                studentrecord.ExamAnswer = "未添加“茶花”";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion
}