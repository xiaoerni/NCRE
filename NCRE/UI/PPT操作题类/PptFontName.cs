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

    #region A卷 查看字体2015年12月11日17:31:56
    public class PptFontNameA
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontNameA()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontName(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "字体";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); 
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides.Count.ToString() == "8")
            {
                studentrecord.ExamAnswer = pp1.Slides[8].Shapes[1].TextFrame.TextRange.Font.Name.ToString();
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

    #region B卷 查看标题字体2015年12月11日17:46:03
    public class PptFontNameTitleB
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontNameTitleB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontName(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "标题字体";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();
           
            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble (dt.Rows[0]["QuestionID"]);;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体

            studentrecord.ExamAnswer = pp1.Slides[1].Shapes[1].TextFrame.TextRange.Font.Name .ToString();
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

    #region B卷 查看字体2015年12月11日17:34:28
    public class PptArtFontNameB
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptArtFontNameB()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontName(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "艺术字字体";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides.Count.ToString() == "8")
            {
                studentrecord.ExamAnswer = pp1.Slides[8].Shapes[1].TextFrame.TextRange.Font.Name.ToString();
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

    #region C卷的查看标题字体2015年12月11日15:41:50
    public class PptFontNameC
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontNameC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontName(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "标题字体";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[1].Shapes[1].TextFrame2.TextRange.Text.ToString() == "禁放烟花爆竹")
            {
                studentrecord.ExamAnswer = pp1.Slides[1].Shapes[1].TextFrame2.TextRange.Font.Name.ToString();
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
                studentrecord.ExamAnswer = "未找到该对象";
                 studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
#endregion

    #region C卷的查看副标题字体2015年12月10日
    public class PptFontSubTitleC
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontSubTitleC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontName(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "字体";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[1].Shapes[2].TextFrame2.TextRange.Text.ToString() == "幼儿看图了解禁放烟花爆竹")
        {
                studentrecord.ExamAnswer = pp1.Slides[1].Shapes[2].TextFrame.TextRange.Font.Name.ToString();
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
                studentrecord.ExamAnswer = "未设置字体";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region C卷的查看动作按钮上的字体名字——隶书
    public class PptButtonNameC
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptButtonNameC()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontName(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "按钮字体";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();
           
            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            //pptinfo.PaperType = dt1.Rows[0]["PaperType"].ToString();

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[5].Shapes.Count > 2)
            {
                studentrecord.ExamAnswer = pp1.Slides[5].Shapes[3].TextFrame.TextRange.Font.Name.ToString();
                if (studentrecord.ExamAnswer == dt.Rows[0]["RightAnswer"].ToString())
                {
                    //studentrecord.ExamAnswer = dt.Rows[0]["RightAnswer"].ToString();
                    studentrecord.Fration = dt.Rows[0]["Fration"].ToString();
                }
                else
                {
                    studentrecord.Fration = "0";

                }
            }
            else
            {
                studentrecord.ExamAnswer = "未添加动作按钮";
                studentrecord.Fration = "0";
            }
            pptquestionbll.ReturnScore(studentrecord);
        }
    }
    #endregion

    #region D卷 查看字体
    public class PptFontNameD
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontNameD()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTD.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontName(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "字体";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            
            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            //pptinfo.PaperType = dt1.Rows[0]["PaperType"].ToString();

            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体

            studentrecord.ExamAnswer = pp1.Slides[5].Shapes[1].TextFrame.TextRange.Font.Name.ToString();
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

    #region H卷 查看字体2015年12月11日17:04:32
    public class PptFontNameH
    {
        Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
        Microsoft.Office.Interop.PowerPoint.Presentation pp1;
        public PptFontNameH()
        {
            pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTH.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }
        private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
        private PptQuestionEntity pptquestion = new PptQuestionEntity();

        public void FontName(PptQuestionEntity pptinfo)
        {
            pptinfo.QuestionFlag = "字体";
            //根据学号查询该学生要考的试题和试卷类型，
            StudentInfoEntity studentinfo = new StudentInfoEntity();

            studentinfo.studentID = FrmLogin.studentID;
            DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
            DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
            PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

            studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]); ;
            studentrecord.StudentID = FrmLogin.studentID;
            //将考生答案传递给studentrecord实体
            if (pp1.Slides[1].Shapes.Count.ToString()=="2")
            {
                studentrecord.ExamAnswer = pp1.Slides[1].Shapes[2].TextFrame.TextRange.Font.Name.ToString();
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
