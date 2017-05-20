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
    #region A卷占位符2015年12月11日17:26:44
    public class PptPlaceholder
    {
         Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptPlaceholder()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void DelPlaceholder(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "删除占位符";

            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();
           
            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble (dt.Rows[0]["QuestionID"]);;
          
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            studentrecord.ExamAnswer = pp1.Slides[1].Shapes.Count.ToString();
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

    #region B卷占位符2015-12-11 17:39:22
    public class PptPlaceholderB
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptPlaceholderB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void DelPlaceholder(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "删除占位符";

            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;

            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            studentrecord.ExamAnswer = pp1.Slides[1].Shapes.Count.ToString();
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
