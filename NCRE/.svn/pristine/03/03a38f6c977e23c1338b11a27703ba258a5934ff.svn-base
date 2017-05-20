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

    #region A卷——第七张幻灯片与第一张幻灯片位置互换2015年12月11日17:27:30
    public class PptMoveA
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;

        public PptMoveA()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void Move(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "移动幻灯片";

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
            studentrecord.ExamAnswer = pp1.Slides[2].Name;
            if (pp1.Slides[2].Name == dt.Rows[0]["RightAnswer"].ToString())
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

    #region B卷——第七张幻灯片与第一张幻灯片位置互换2015年12月11日17:40:29
    public class PptMoveB
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;

        public PptMoveB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void Move(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "移动幻灯片";

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
            studentrecord.ExamAnswer = pp1.Slides[2].Name;
            if (pp1.Slides[2].Name == dt.Rows[0]["RightAnswer"].ToString())
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

    #region C卷——第三和第一张交换位置2015年12月11日15:35:30
    public class PptMoveC
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;

        public PptMoveC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void Move(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "移动幻灯片";

            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();
           
            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble (dt.Rows[0]["QuestionID"]);
            //传递考生ID进studentrecord实体
            studentrecord.StudentID = FrmLogin.studentID;
            //学生答案
            studentrecord.ExamAnswer = pp1.Slides[1].Name.ToString();
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[1].Name == dt.Rows[0]["RightAnswer"].ToString())
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


    #region E卷——第三和第四张交换位置2015年12月11日16:34:11
    public class PptMoveE
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;

        public PptMoveE()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTE.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void Move(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "移动幻灯片";

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
           studentrecord.ExamAnswer = pp1.Slides[3].Name.ToString();
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

    #region F卷——第三和第四张交换位置2015年12月11日16:41:47
    public class PptMoveF
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;

        public PptMoveF()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTF.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void Move(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "移动幻灯片";

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
            studentrecord.ExamAnswer = pp1.Slides[3].Name.ToString();
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
