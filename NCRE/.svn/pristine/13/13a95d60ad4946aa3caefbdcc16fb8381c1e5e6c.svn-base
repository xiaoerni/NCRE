using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using Model;
using MSExcel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.IO;
using System.Reflection;
using System.Data.SqlClient;
using System.Data;

namespace NCRE学生考试端V1._0
{
    public class ExcelRowHeight
    {
        private ExcelEntityBLL excelquestionblrowheight = new ExcelEntityBLL();
        #region 指定位置单元格行高——王虹芸

        public void RowHeight(ExcelQuestionEntity excelinfo)
        {
            //将正确答案、分值取出来，传给dt
            excelinfo.QuestionFlag = "行高";
            DataTable dt = excelquestionblrowheight.QueryExcelTypeID(excelinfo);
            ExcelQuestionRecordEntity excelrecord = new ExcelQuestionRecordEntity();
            string fration;
            //传递考生ID进excelrecord实体
            excelrecord.StudentID = MyInfo.MystudentID();
            //获取试卷类型
            excelrecord.PaperType = MyInfo.MyPaperType();
            excelrecord.ExamAnswer = "考生未答题";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传入试卷试题ID
                excelrecord.QuestionID = Convert.ToDouble(excelinfo.QuestionTypeID);

                //获取试卷试题内容
                excelrecord.QuestionContent = dt.Rows[i]["QuestionContent"].ToString();
                //将正确答案传给excelrecord实体
                excelrecord.CorrectAnswer = dt.Rows[i]["CorrectAnswer"].ToString();
                //将实体的QuestionID选择出来
                excelrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //获取指定单元格位置
                string x = dt.Rows[i]["PositionX"].ToString();
                try
                {
                    //获取工作表
                    MSExcel.Worksheet sheet1 = ExcelJudgeHelper.m_workbook.ActiveSheet as MSExcel.Worksheet;

                    string rowheight = ExcelJudgeHelper.m_excel.Range[x].RowHeight.ToString();
                    excelrecord.ExamAnswer = rowheight;
                }
                catch { }
                if (excelrecord.ExamAnswer == excelrecord.CorrectAnswer)
                {
                    excelrecord.Fration = Convert.ToDouble(fration);
                }
                else
                {
                    excelrecord.Fration = 0;
                }
                excelquestionblrowheight.ReturnExcelScore(excelrecord);
            }
        }
        #endregion
    }
}
