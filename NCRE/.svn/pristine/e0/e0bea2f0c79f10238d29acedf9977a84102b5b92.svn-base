
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using MSExcel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Threading;
using System.Runtime.InteropServices;//互动服务 
using IWshRuntimeLibrary;


namespace NCRE学生考试端V1._0
{
    public class ExcelFilesChazhao    
    {
        private ExcelEntityBLL excelquestionbllfileschaozhao = new ExcelEntityBLL();
    
        #region"查找文件夹-王虹芸"
        /// <summary>
        /// 查找文件夹
        /// </summary>
        /// <param name="winquestion"></param>
        public void FilesChaozhao(ExcelQuestionEntity excelinfo)
        {          
            //将正确答案，分值取出来，传给studentRecord
            excelinfo.QuestionFlag = "文件夹查找";
            DataTable dt = excelquestionbllfileschaozhao.QueryExcelTypeID(excelinfo);
            ExcelQuestionRecordEntity excelrecord = new ExcelQuestionRecordEntity();
            //将考生ID传到studentRecord实体
            excelrecord.StudentID = MyInfo.MystudentID();
            excelrecord.PaperType = MyInfo.MyPaperType();
            string correctAnswer;
            string fraction;
            string examAnswer;
            //循环遍历正确答案
            excelrecord.ExamAnswer = "考生未答题";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //将试题的ID取出来
                excelrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["questionID"]);
                //将题的分数取出来
                fraction = dt.Rows[i]["Fration"].ToString();
                //将正确答案取出来
                correctAnswer = dt.Rows[i]["CorrectAnswer"].ToString();

                string str = @"D:\计算机一级考生文件\Excelkt\" + correctAnswer;

                if (System.IO.File.Exists(str))
                {
                    //加分
                    excelrecord.Fration = Convert.ToDouble(fraction);
                    examAnswer = correctAnswer;
                    excelrecord.ExamAnswer = examAnswer;
                }
                else
                {
                    //不加分
                    excelrecord.Fration = 0;
                }
                excelquestionbllfileschaozhao.ReturnExcelScore(excelrecord);
            }
        }
        #endregion
        
    }
}



