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
    public class ExcelChartColor
    {
        private ExcelEntityBLL excelquestionbll3 = new ExcelEntityBLL();

        #region excel根据图表颜色判分——王虹芸
        /// <summary>
        /// 图表题判分
        /// </summary>
        /// <param name="excelinfo"></param>
        public void ChartColor(ExcelQuestionEntity excelinfo)
        {
            //将正确答案、分值取出来，传给dt
            excelinfo.QuestionFlag = "图表颜色";
            System.Data.DataTable dt = excelquestionbll3.QueryExcelTypeID(excelinfo);
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
                excelrecord.QuestionID = Convert.ToDouble (excelinfo.QuestionTypeID);
                              
                //获取试卷试题内容
                excelrecord.QuestionContent = dt.Rows[i]["QuestionContent"].ToString();

                //将正确答案传给excelrecord实体
                excelrecord.CorrectAnswer = dt.Rows[i]["CorrectAnswer"].ToString();
                //将实体的QuestionID选择出来
                excelrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]);
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
                    //获取图表颜色
                    excelrecord.ExamAnswer = chartobject1.Chart.ChartArea.Interior.ColorIndex.ToString();
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
                excelquestionbll3.ReturnExcelScore(excelrecord);
            }
        }
        #endregion
    }
}
