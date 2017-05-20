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
#region D卷幻灯片声音2015年12月11日16:21:58
   public  class PptSoundD
    {
       Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptSoundD()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTD.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void actionType(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "声音";
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

            if (pp1.Slides[1].SlideShowTransition.SoundEffect.Type.ToString() == "ppSoundNone")
            {
               studentrecord.ExamAnswer = "ppSoundNone";
                studentrecord.Fration = "0";
            }
            else
            {
                studentrecord.ExamAnswer = pp1.Slides[1].SlideShowTransition.SoundEffect.Name;

                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
                {

                    studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
                }
                else
                {
                    studentrecord.Fration = "0";
                }
            }
            
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

   #region F卷幻灯片声音2015年12月11日16:44:16
   public class PptSoundF
   {
       Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
       Microsoft.Office.Interop.PowerPoint.Presentation pp1;
       public PptSoundF()
       {
           pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTF.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

       }
       private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
       private PptQuestionEntity pptquestion = new PptQuestionEntity();

       public void actionType(PptQuestionEntity pptinfo)
       {
           pptinfo.QuestionFlag = "声音";
           //根据学号查询该学生要考的试题和试卷类型，
           StudentInfoEntity studentinfo = new StudentInfoEntity();

           studentinfo.studentID = FrmLogin.studentID;
           DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
           DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
           PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
           studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
           //传递考生ID进studentrecord实体
           studentrecord.StudentID = FrmLogin.studentID;
           if (pp1.Slides[1].SlideShowTransition.SoundEffect.Type.ToString() == "ppSoundNone")
           {
               studentrecord.ExamAnswer = "ppSoundNone";
               studentrecord.Fration = "0";
           }
           else {
           
           //将考生答案传递给studentrecord实体
               
                   studentrecord.ExamAnswer = pp1.Slides[1].SlideShowTransition.SoundEffect.Name;

                   if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
                   {

                       studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
                   }

                   else
                   {
                       studentrecord.Fration = "0";
                   }
               }

           
           pptquestionbll.ReturnScore(studentrecord);
       }
   }
   #endregion

}

