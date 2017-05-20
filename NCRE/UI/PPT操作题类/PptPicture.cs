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
    public class PptPicture
    {
        #region D卷插入图片——查看图片是否正确
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptPicture()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTD.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void pptpicture(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "插入图片";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            
            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
           // pptinfo.PaperType = dt1.Rows[0]["PaperType"].ToString();

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble (dt.Rows[0]["QuestionID"]);
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[7].Shapes.Count > 1)
            {
                studentrecord.ExamAnswer = pp1.Slides[7].Shapes[2].Name.ToString();
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
                studentrecord.ExamAnswer = "未添加图片";
                studentrecord.Fration = "0";

            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
        #endregion

    #region D卷插入图片——查看图片大小2015年12月11日16:05:22
    public class PptPictureSizeD
    {
    
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptPictureSizeD()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTD.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void pptpicturesize(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "图片大小";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            
            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
             DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble (dt.Rows[0]["QuestionID"]);
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体

            if (pp1.Slides[3].Shapes.Count == 5)
            {
                studentrecord.ExamAnswer = pp1.Slides[3].Shapes[5].Height.ToString();
            
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
            studentrecord .ExamAnswer ="未添加图片";
            studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
        #endregion

    #region C卷插入图片——查看图片大小2015年12月10日
    public class PptPictureSizeC
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptPictureSizeC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void pptpicturesize(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "图片大小";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体

            if (pp1.Slides[2].Shapes.Count == 3) 
            {
                studentrecord.ExamAnswer = pp1.Slides[2].Shapes[2].Width.ToString();

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
                studentrecord.ExamAnswer = "未添加图片";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region F卷插入图片——查看图片大小2015年12月11日16:45:40
    public class PptPictureSizeF
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptPictureSizeF()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTF.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void pptpicturesize(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "图片大小";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体


                if (pp1.Slides[4].Shapes.Count.ToString()=="4")
                {
                studentrecord.ExamAnswer = pp1.Slides[4].Shapes[4].Width.ToString();

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
                     studentrecord.ExamAnswer ="长和宽设置不同";
                         studentrecord.Fration = "0";
                    }
           
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region H卷插入图片——查看图片大小宽：2015年12月11日17:13:59
    public class  PptTextWidthH
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptTextWidthH()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTH.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void PptTextWidth(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "图片大小";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            studentrecord.ExamAnswer = pp1.Slides[4].Shapes[4].Width.ToString();


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

    #region H卷插入图片——查看图片大小高：2015年12月11日17:13:14
    public class PptPictureSizeH
    {

        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptPictureSizeH()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTH.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void pptpicturesize(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "插入图片";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            studentrecord.ExamAnswer = pp1.Slides[4].Shapes[4].Height .ToString();

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
