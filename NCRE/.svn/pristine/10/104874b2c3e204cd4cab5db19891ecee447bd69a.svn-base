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

namespace NCRE学生考试端V1._0
{
    public class ExcelChartPosition
    {
        private ExcelEntityBLL excelquestionbllchartposition = new ExcelEntityBLL();

        #region excel根据图表位置判分——王虹芸
        /// <summary>
        /// 图表位置判分
        /// </summary>
        /// <param name="excelinfo"></param>
        public void ChartPosition(ExcelQuestionEntity excelinfo)
        {
            //将正确答案、分值取出来，传给dt
            excelinfo.QuestionFlag = "图表位置";
            System.Data.DataTable dt = excelquestionbllchartposition.QueryExcelTypeID(excelinfo);
            ExcelQuestionRecordEntity excelrecord = new ExcelQuestionRecordEntity();
            string fration;
            excelrecord.ExamAnswer = "考生未答题";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传入试卷试题ID
                excelrecord.QuestionID = Convert.ToDouble(excelinfo.QuestionTypeID);
                //传递考生ID进excelrecord实体
                excelrecord.StudentID = MyInfo.MystudentID();
                //获取试卷正确答案

                //获取试卷类型
                excelrecord.PaperType = MyInfo.MyPaperType();
                //获取试卷试题内容
                excelrecord.QuestionContent = dt.Rows[i]["QuestionContent"].ToString();
                //将正确答案传给excelrecord实体
                excelrecord.CorrectAnswer = dt.Rows[i]["CorrectAnswer"].ToString();
                //将实体的QuestionID选择出来
                excelrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]); ;

                //将每道题分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //获取图表所在工作表位置
                string x = dt.Rows[i]["PositionX"].ToString();

                int intX = int.Parse(x);
                try
                {
                    //获取工作表及图表
                    MSExcel.Worksheet sheet1 = (Worksheet)ExcelJudgeHelper.m_workbook.Worksheets[intX];
                    MSExcel.ChartObject chartobject1 = (MSExcel.ChartObject)sheet1.ChartObjects("图表 1");
                    //获取图表位置               
                    string HorizonX = chartobject1.TopLeftCell.Row.ToString();
                    excelrecord.ExamAnswer = HorizonX;
                }
                catch { }
                if (excelrecord.CorrectAnswer == excelrecord.ExamAnswer)
                {
                    excelrecord.Fration = Convert.ToDouble(fration);
                }
                else
                {
                    excelrecord.Fration = 0;
                }
                excelquestionbllchartposition.ReturnExcelScore(excelrecord);
            }
        }
        #endregion
    }
}
