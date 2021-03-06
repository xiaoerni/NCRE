﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using Model;
using System.Data;
using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;

//这个类是负责将word中的查找替换题来做的,各个试卷都可以用

namespace NCRE学生考试端V1._0
{
    public class WordAFindKeyWord
    {
        private static  Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static  object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordA.docx"; 
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();

       
        #region 修改人 王荣晓 查找替换
        /// <summary>
        /// 查找替换
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindKeyWord(WordQuestionEntity wordinfo)
        {
           
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "插入空行";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Text.Trim().Substring(0, 7);           
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Paragraphs[2].Range.Text.Trim().Substring(0, 7) == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
                {
                    //加分
                    studentrecord.Fration = fration;
                }
                else
                {
                    //不加分
                    studentrecord.Fration = "0";
                }
                wordquestionbll.ReturnScore(studentrecord);
            }
            return;
        }
        #endregion

        #region 修改人 王荣晓 查找替换
        /// <summary>
        /// 查找替换
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindKeyWordA(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找替换";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Content.Text.Trim().Substring(35, 3);
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Content.Text.Trim().Substring(35, 3) == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
                {
                    //加分
                    studentrecord.Fration = fration;
                }
                else
                {
                    //不加分
                    studentrecord.Fration = "0";
                }
                wordquestionbll.ReturnScore(studentrecord);
            }
            return;
        }
        #endregion
        

     }
}
