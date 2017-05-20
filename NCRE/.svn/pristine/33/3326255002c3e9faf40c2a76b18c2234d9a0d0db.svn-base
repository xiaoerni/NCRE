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
    #region A卷查看标题的加粗2015年12月11日17:32:28
    public class PptBoldA
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptBoldA()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void Bold(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "加粗";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体


            if (pp1.Slides.Count.ToString() == "8")
            {
                studentrecord.ExamAnswer = pp1.Slides[8].Shapes[1].TextFrame2.TextRange.Font.Bold.ToString();
                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString().Trim())
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

    #region B卷查看标题的加粗2015年12月11日17:47:24
    public class PptBoldB
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptBoldB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void bold(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "加粗";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体


            if (pp1.Slides.Count.ToString() == "8")
            {
                studentrecord.ExamAnswer = pp1.Slides[8].Shapes[1].TextFrame.TextRange.Font.Bold.ToString();
                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString().Trim())
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

    #region C卷查看主标题的加粗2015年12月11日15:41:22
    public class PptBoldC
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptBoldC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void BoldC(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "加粗";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体


            if (pp1.Slides[1].Shapes[1].TextFrame2.TextRange.Text.ToString() == "禁放烟花爆竹")
            {
                studentrecord.ExamAnswer = pp1.Slides[1].Shapes[1].TextFrame.TextRange.Font.Bold.ToString();
                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString().Trim())
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

                studentrecord.ExamAnswer = "未找到该对象";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
       
    }
    
    #endregion
}
