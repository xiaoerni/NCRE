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
    public class ExcelSheetPosition
    {
        private ExcelEntityBLL excelquestionblsheetposition = new ExcelEntityBLL();

        #region excel查找工作表内容位置——王虹芸
        /// <summary>
        /// 查找内容位置    
        /// </summary>
        /// <param name="excelinfo"></param>
        public void SheetPosition(ExcelQuestionEntity excelinfo)
        {
            //将正确答案、分值取出来，传给dt
            excelinfo.QuestionFlag = "查找工作表内容位置";
            System.Data.DataTable dt = excelquestionblsheetposition.QueryExcelTypeID(excelinfo);
            ExcelQuestionRecordEntity excelrecord = new ExcelQuestionRecordEntity();
            string fration;

            //将考生ID传到studentRecord实体
            excelrecord.StudentID = MyInfo.MystudentID();
            //获取试卷类型
            excelrecord.PaperType = MyInfo.MyPaperType();
            excelrecord.ExamAnswer = "考生未答题";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //获取试卷试题内容
                excelrecord.QuestionContent = dt.Rows[i]["QuestionContent"].ToString();
                //将正确答案传给excelrecord实体
                excelrecord.CorrectAnswer = dt.Rows[i]["CorrectAnswer"].ToString();
                //将实体的QuestionID选择出来
                excelrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);

                //将每道题分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //获取查找内容位置
                string x = dt.Rows[i]["PositionX"].ToString ();
                string y = dt.Rows[i]["PositionY"].ToString().Trim();

                int intX = int.Parse(x);
                try
                {
                    //获取工作表
                    Worksheet sheet1 = (Worksheet)ExcelJudgeHelper.m_workbook.Worksheets[intX];
                    string[] sArray = Regex.Split(y, ",", RegexOptions.IgnoreCase);
                    int firstNum = int.Parse(sArray[0]);
                    int secondNum = int.Parse(sArray[1]);

                    Range currentCell = (Range)sheet1.Cells[firstNum, secondNum];

                    excelrecord.ExamAnswer = currentCell.Text.ToString();
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
                excelquestionblsheetposition.ReturnExcelScore(excelrecord);
            }
        }
        #endregion
    }
}
