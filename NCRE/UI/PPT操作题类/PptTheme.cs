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

namespace NCRE学生考试端V1._0.PPT操作题类
{

    #region A卷的主题设置2015年12月11日17:33:41
    public class PptThemeA
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptThemeA()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void Theme(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "主题";
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
            studentrecord.ExamAnswer = pp1.Slides[1].Design.Name.ToString();

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

    #region B卷的主题设置2015年12月11日17:48:23
    public class PptThemeB
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptThemeB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void themeB(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "主题";
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
            studentrecord.ExamAnswer = pp1.Slides[1].Design.Name.ToString();

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

    #region C卷的主题设置2015年12月11日15:35:15
    public class PptThemeC
    {
          Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptThemeC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "主题";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            
            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble (dt.Rows[0]["QuestionID"]);;
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            studentrecord.ExamAnswer = pp1.Slides[1].Design.Name.ToString();

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

    #region E卷的主题设置2015年12月11日16:35:29
    public class PptThemeE
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptThemeE()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTE.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "主题";
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
            studentrecord.ExamAnswer = pp1.Slides[1].Design.Name.ToString();

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
