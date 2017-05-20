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


namespace NCRE学生考试端V1._0
{
    #region A卷切换效果2015年12月11日17:33:08
    public class PptActionTypeA
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptActionTypeA()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "切换效果";
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
            studentrecord.ExamAnswer = pp1.Slides[1].SlideShowTransition.EntryEffect.ToString();

            if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
            {

                studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
            }
            else
            {
                studentrecord.Fration = "0";
            }

            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region B卷切换效果2015年12月11日17:47:48
    public class PptActionTypeB
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptActionTypeB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "切换效果";
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
            studentrecord.ExamAnswer = pp1.Slides[1].SlideShowTransition.EntryEffect.ToString();

            if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
            {

                studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
            }
            else
            {
                studentrecord.Fration = "0";
            }

            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region C卷切换效果2015年12月11日15:51:10
    public class PptActionTypeC
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptActionTypeC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "切换效果";
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
            studentrecord.ExamAnswer = pp1.Slides[1].SlideShowTransition.EntryEffect.ToString();

            if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
            {

                studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
            }
            else
            {
                studentrecord.Fration = "0";
            }

            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region D卷切换效果：百叶窗垂直2015年12月11日16:17:03
    public class PptActionTypeD
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptActionTypeD()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTD.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "切换效果";
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
            studentrecord.ExamAnswer = pp1.Slides[1].SlideShowTransition.EntryEffect.ToString();

            if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
            {

                studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
            }
            else
            {
                studentrecord.Fration = "0";
            }

            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region E卷切换效果2015年12月11日16:33:11
    public class PptActionTypeE
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptActionTypeE()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTE.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "切换效果";
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
            studentrecord.ExamAnswer = pp1.Slides[2].SlideShowTransition.EntryEffect.ToString();

            if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
            {

                studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
            }
            else
            {
                studentrecord.Fration = "0";
            }

            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region F卷切换效果2015年12月11日16:44:22
    public class PptActionTypeF
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptActionTypeF()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTF.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "切换效果";
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
            studentrecord.ExamAnswer = pp1.Slides[1].SlideShowTransition.EntryEffect.ToString();

            if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
            {

                studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
            }
            else
            {
                studentrecord.Fration = "0";
            }

            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region G卷切换效果2015-12-11 16:58:26
    public class PptActionTypeG
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptActionTypeG()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTG.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "动画效果";
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
            studentrecord.ExamAnswer = pp1.Slides[2].SlideShowTransition.EntryEffect.ToString();

            if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
            {

                studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
            }
            else
            {
                studentrecord.Fration = "0";
            }

            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region H卷切换效果2015年12月11日17:06:47
    public class PptActionTypeH
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptActionTypeH()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTH.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "切换效果";
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
            studentrecord.ExamAnswer = pp1.Slides[1].SlideShowTransition.EntryEffect.ToString();

            if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
            {

                studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
            }
            else
            {
                studentrecord.Fration = "0";
            }

            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

}

