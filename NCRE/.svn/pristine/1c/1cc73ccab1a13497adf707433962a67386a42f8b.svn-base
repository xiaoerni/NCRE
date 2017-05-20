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

//艺术字插入字
namespace NCRE学生考试端V1._0.PPT操作题类
{

        #region A卷艺术字文本2015年12月11日17:31:17
        public class PptArtWordTextA
        {
            Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            Microsoft.Office.Interop.PowerPoint.Presentation pp1;
            public PptArtWordTextA()
            {
                pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

            }
            private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
            private PptQuestionEntity pptquestion = new PptQuestionEntity();

            public void WordtextA(PptQuestionEntity pptinfo)
            {
                pptinfo.QuestionFlag = "插入字";
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
                    studentrecord.ExamAnswer = pp1.Slides[8].Shapes[1].TextFrame2.TextRange.Text.ToString();
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

        #region B卷艺术字文本2015年12月11日17:43:22
        public class PptArtWordTextB
        {
            Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            Microsoft.Office.Interop.PowerPoint.Presentation pp1;
            public PptArtWordTextB()
            {
                pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

            }
            private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
            private PptQuestionEntity pptquestion = new PptQuestionEntity();

            public void Artword(PptQuestionEntity pptinfo)
            {
                pptinfo.QuestionFlag = "插入字";
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
                    studentrecord.ExamAnswer = pp1.Slides[8].Shapes[1].TextFrame2.TextRange.Text.ToString();
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

        #region E卷艺术字文本2015年12月11日16:35:16
        public class PptArtWordTextE
        {
            Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            Microsoft.Office.Interop.PowerPoint.Presentation pp1;
            public PptArtWordTextE()
            {
                pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTE.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

            }
            private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
            private PptQuestionEntity pptquestion = new PptQuestionEntity();

            public void actionType(PptQuestionEntity pptinfo)
            {
                pptinfo.QuestionFlag = "艺术字文本";
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
                    studentrecord.ExamAnswer = pp1.Slides[5].Shapes[1].TextFrame2.TextRange.Text.ToString();
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

        #region H卷艺术字文本2015年12月11日17:02:46
        public class PptArtWordTextH
        {
            Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            Microsoft.Office.Interop.PowerPoint.Presentation pp1;
            public PptArtWordTextH()
            {
                pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTH.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

            }
            private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
            private PptQuestionEntity pptquestion = new PptQuestionEntity();

            public void actionType(PptQuestionEntity pptinfo)
            {
                pptinfo.QuestionFlag = "艺术字文本";
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
                
                if (pp1.Slides[1].Shapes.Count.ToString()=="2")
                {
                    studentrecord.ExamAnswer = pp1.Slides[1].Shapes[2].TextFrame2.TextRange.Text.ToString();
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

        #region H卷文本框2015年12月11日17:09:56
        public class PptTextSizeH
        {
            Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            Microsoft.Office.Interop.PowerPoint.Presentation pp1;
            public PptTextSizeH()
            {
                pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTH.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

            }
            private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
            private PptQuestionEntity pptquestion = new PptQuestionEntity();

            public void TextH(PptQuestionEntity pptinfo)
            {
                pptinfo.QuestionFlag = "文本框插入字";
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
               
                if (pp1.Slides.Count.ToString() == "3")
                {
                    studentrecord.ExamAnswer = pp1.Slides[3].Shapes[3].TextFrame2.TextRange.Text.ToString();
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
                    studentrecord.ExamAnswer = "未添加文本框";
                    studentrecord.Fration = "0";
                }
                pptquestionbll.ReturnScore(studentrecord);
            }
        }
        #endregion
    }


