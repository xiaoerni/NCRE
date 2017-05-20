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
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace NCRE学生考试端V1._0
{
    public class ExcelVerticalAlignment
    {
        private ExcelEntityBLL excelquestionblverticalalignment = new ExcelEntityBLL();

        #region 单元格垂直对齐方式——王虹芸
        public void VerticalAlinment(ExcelQuestionEntity excelinfo)
        {
            //将正确答案、分值取出来，传给dt
            excelinfo.QuestionFlag = "单元格垂直对齐方式";
            System.Data.DataTable dt = excelquestionblverticalalignment.QueryExcelTypeID(excelinfo);
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

                excelrecord.QuestionContent = dt.Rows[i]["QuestionContent"].ToString();

                //将正确答案传给excelrecord实体
                excelrecord.CorrectAnswer = dt.Rows[i]["CorrectAnswer"].ToString();
                //将实体的QuestionID选择出来
                excelrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                string x = dt.Rows[i]["PositionX"].ToString();
                string y = dt.Rows[i]["PositionY"].ToString();
                int intX = int.Parse(x);
                try
                {
                    //获取工作表
                    Worksheet sheet1 = (Worksheet)ExcelJudgeHelper.m_workbook.Worksheets[intX];
                    string[] sArray = Regex.Split(y, ",", RegexOptions.IgnoreCase);
                    string firstNum = sArray[0];
                    string secondNum = sArray[1];
                    excelrecord.ExamAnswer = ExcelJudgeHelper.m_excel.Range[firstNum, secondNum].VerticalAlignment.ToString();
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
                excelquestionblverticalalignment.ReturnExcelScore(excelrecord);
            }
        }
        #endregion
    }
}
