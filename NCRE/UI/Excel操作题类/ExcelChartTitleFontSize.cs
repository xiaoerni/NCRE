﻿using System;
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
    public class ExcelChartTitleFontSize
    {
        private ExcelEntityBLL excelquestionbllcharttitlefontsize = new ExcelEntityBLL();

        #region excel根据图表标题大小判分——王虹芸
        /// <summary>
        /// 图表题判分
        /// </summary>
        /// <param name="excelinfo"></param>
        public void ChartTitleFontSize(ExcelQuestionEntity excelinfo)
        {
            //将正确答案、分值取出来，传给dt
            excelinfo.QuestionFlag = "图表标题大小";
            System.Data.DataTable dt = excelquestionbllcharttitlefontsize.QueryExcelTypeID(excelinfo);
            ExcelQuestionRecordEntity excelrecord = new ExcelQuestionRecordEntity();
            string a;
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

                //获取试卷正确答案
                a = dt.Rows[i]["CorrectAnswer"].ToString();

                //获取试卷试题内容
                excelrecord.QuestionContent = dt.Rows[i]["QuestionContent"].ToString();
                //将正确答案传给excelrecord实体
                excelrecord.CorrectAnswer = a.ToString();
                //将实体的QuestionID选择出来
                excelrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
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
                    //获取图表标题大小
                    string charttitlesize = chartobject1.Chart.ChartTitle.Font.Size.ToString();
                    excelrecord.ExamAnswer = charttitlesize;
                }
                catch (Exception)
                { }
                if (a == excelrecord.ExamAnswer)
                {
                    excelrecord.Fration = Convert.ToDouble(fration);
                }
                else
                {
                    excelrecord.Fration = 0;
                }
                excelquestionbllcharttitlefontsize.ReturnExcelScore(excelrecord);
            }
        }
        #endregion
    }
}
