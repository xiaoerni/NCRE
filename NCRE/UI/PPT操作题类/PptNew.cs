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
using NCRE学生考试端V1._0;


namespace NCRE学生考试端V1._0
{
#region A卷添加幻灯片2015年12月11日17:29:07
public class PptNewA
{

    Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
    Microsoft.Office.Interop.PowerPoint.Presentation pp1;

    public PptNewA()
    {
        pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTA.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
    }
    private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
    private PptQuestionEntity pptquestion = new PptQuestionEntity();

    public void New(PptQuestionEntity pptinfo)
    {
        pptinfo.QuestionFlag = "添加幻灯片";

        //根据学号查询该学生要考的试题和试卷类型，
        StudentInfoEntity studentinfo = new StudentInfoEntity();

        studentinfo.studentID = FrmLogin.studentID;
        DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
        DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
        PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

        studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
        //传递考生ID进studentrecord实体
        studentrecord.StudentID = FrmLogin.studentID;
        //学生答案
        studentrecord.ExamAnswer = pp1.Slides.Count.ToString();
        //将考生答案传递给studentrecord实体
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

#region B卷添加幻灯片2015年12月11日17:42:39
public class PptNewB
{

    Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
    Microsoft.Office.Interop.PowerPoint.Presentation pp1;

    public PptNewB()
    {
        pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTB.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
    }
    private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
    private PptQuestionEntity pptquestion = new PptQuestionEntity();

    public void Add(PptQuestionEntity pptinfo)
    {
        pptinfo.QuestionFlag = "添加幻灯片";

        //根据学号查询该学生要考的试题和试卷类型，
        StudentInfoEntity studentinfo = new StudentInfoEntity();

        studentinfo.studentID = FrmLogin.studentID;
        DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
        DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
        PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

        studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
        //传递考生ID进studentrecord实体
        studentrecord.StudentID = FrmLogin.studentID;
        //学生答案
        studentrecord.ExamAnswer = pp1.Slides.Count.ToString();
        //将考生答案传递给studentrecord实体
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

#region C卷删除幻灯片2015年12月11日15:35:46
public class PptDelC
{

    Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
    Microsoft.Office.Interop.PowerPoint.Presentation pp1;

    public PptDelC()
    {
        pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTC.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
    }
    private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
    private PptQuestionEntity pptquestion = new PptQuestionEntity();

    public void Move(PptQuestionEntity pptinfo)
    {
        pptinfo.QuestionFlag = "删除幻灯片";

        //根据学号查询该学生要考的试题和试卷类型，
        StudentInfoEntity studentinfo = new StudentInfoEntity();

        studentinfo.studentID = FrmLogin.studentID;
        DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
        DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
        PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

        studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
        //传递考生ID进studentrecord实体
        studentrecord.StudentID = FrmLogin.studentID;
        //学生答案
        studentrecord.ExamAnswer = pp1.Slides.Count.ToString();
        //将考生答案传递给studentrecord实体
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

#region D卷插入幻灯片2015年12月11日16:04:52
public class PptNewD
{

    Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
    Microsoft.Office.Interop.PowerPoint.Presentation pp1;

    public PptNewD()
    {
        pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTD.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
    }
    private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
    private PptQuestionEntity pptquestion = new PptQuestionEntity();

    public void Move(PptQuestionEntity pptinfo)
    {
        pptinfo.QuestionFlag = "插入幻灯片";

        //根据学号查询该学生要考的试题和试卷类型，
        StudentInfoEntity studentinfo = new StudentInfoEntity();

        studentinfo.studentID = FrmLogin.studentID;
        DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
        DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
        PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

        studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
        //传递考生ID进studentrecord实体
        studentrecord.StudentID = FrmLogin.studentID;
        //学生答案
        studentrecord.ExamAnswer = pp1.Slides.Count.ToString();
        //将考生答案传递给studentrecord实体
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

#region E卷添加幻灯片2015年12月11日16:35:23
public class PptNewE
{

    Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
    Microsoft.Office.Interop.PowerPoint.Presentation pp1;

    public PptNewE()
    {
        pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTE.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
    }
    private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
    private PptQuestionEntity pptquestion = new PptQuestionEntity();

    public void Move(PptQuestionEntity pptinfo)
    {
        pptinfo.QuestionFlag = "添加幻灯片";

        //根据学号查询该学生要考的试题和试卷类型，
        StudentInfoEntity studentinfo = new StudentInfoEntity();

        studentinfo.studentID = FrmLogin.studentID;
        DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
        DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
        PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

        studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
        //传递考生ID进studentrecord实体
        studentrecord.StudentID = FrmLogin.studentID;
        //学生答案
        studentrecord.ExamAnswer = pp1.Slides.Count.ToString();
        //将考生答案传递给studentrecord实体
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

#region G卷添加幻灯片2015年12月11日16:57:00
public class PptNewG
{

    Microsoft.Office.Interop.PowerPoint.Application pa1 = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
    Microsoft.Office.Interop.PowerPoint.Presentation pp1;

    public PptNewG()
    {
        pp1 = pa1.Presentations.Open(@"D:\计算机一级考生文件\Pptkt\PPTG.pptx", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
    }
    private PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();
    private PptQuestionEntity pptquestion = new PptQuestionEntity();

    public void Move(PptQuestionEntity pptinfo)
    {
        pptinfo.QuestionFlag = "添加幻灯片";

        //根据学号查询该学生要考的试题和试卷类型，
        StudentInfoEntity studentinfo = new StudentInfoEntity();

        studentinfo.studentID = FrmLogin.studentID;
        DataTable dt1 = pptquestionbll.SelectPaperTypeByStudentIDBLL(studentinfo);
        DataTable dt = pptquestionbll.LoadPptByFlag(pptinfo);
        PptQuestionRecordEntity studentrecord = new PptQuestionRecordEntity();

        studentrecord.QuestionID = Convert.ToDouble(dt.Rows[0]["QuestionID"]);
        //传递考生ID进studentrecord实体
        studentrecord.StudentID = FrmLogin.studentID;
        //学生答案
        studentrecord.ExamAnswer = pp1.Slides.Count.ToString();
        //将考生答案传递给studentrecord实体
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