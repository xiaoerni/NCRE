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
    #region A卷字体颜色2015年12月11日17:20:16
    public class PptFontColor
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontColor()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontColor(PptQuestionEntity pptinfo)
        {
            
           
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = FrmLogin.studentID;
            pptinfo.QuestionFlag = "字体颜色";
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
             PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
             studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;



            if (pp1.Slides[1].Shapes[1].TextFrame.TextRange.Text == "战 斗 机")//如果是文字,文字处理
            {
                //将考生答案传递给studentrecord实体
                Microsoft.Office.Interop.PowerPoint.ColorFormat mycolor = pp1.Slides[1].Shapes[1].TextFrame.TextRange.Font.Color;
                studentrecord.ExamAnswer = mycolor.RGB.ToString();

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
                studentrecord.ExamAnswer = "未设置颜色";
                studentrecord.Fration = "0";
            }
                pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region B卷字体颜色2015年12月11日17:20:16
    public class PptFontColorB
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontColorB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontColor(PptQuestionEntity pptinfo)
        {


            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = FrmLogin.studentID;
            pptinfo.QuestionFlag = "字体颜色";
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();
            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;



            if (pp1.Slides[1].Shapes[1].TextFrame.TextRange.Text == "战 斗 机")//如果是文字,文字处理
            {
                //将考生答案传递给studentrecord实体
                Microsoft.Office.Interop.PowerPoint.ColorFormat mycolor = pp1.Slides[1].Shapes[1].TextFrame.TextRange.Font.Color;
                studentrecord.ExamAnswer = mycolor.RGB.ToString();

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
                studentrecord.ExamAnswer = "未设置颜色";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region C卷字体颜色
    public class PptFontColorC
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontColorC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontColor(PptQuestionEntity pptinfo)
        {
          
                pptinfo.QuestionFlag = "字体颜色";

                //根据学号查询该学生要考的试题和试卷类型，
                StudentInfoEntity studentinfo = new StudentInfoEntity();
                
                studentinfo.studentID = FrmLogin.studentID;
                DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
                //pptinfo.PaperType = dt1.Rows[0]["PaperType"].ToString();

                DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
                PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体

                    studentrecord.StudentID = FrmLogin.studentID;
                    //将考生答案传递给studentrecord实体
                    Microsoft.Office.Interop.PowerPoint.ColorFormat mycolor = pp1.Slides[1].Shapes[1].TextFrame.TextRange.Font.Color;
                    studentrecord.ExamAnswer = mycolor.RGB.ToString();
                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    if (studentrecord.ExamAnswer == "153")
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
        }
    
    #endregion
}
