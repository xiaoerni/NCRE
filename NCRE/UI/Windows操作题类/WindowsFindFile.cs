﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Threading;

namespace NCRE学生考试端V1._0
{
    public class WindowsFindFile
    {
        private WinQuestionEntityBLL winquestionbll = new WinQuestionEntityBLL();

        #region"查找文件-韩梦甜修改-2015-11-20"
        /// <summary>
        /// 查找文件
        /// </summary>
        /// <param name="winquestion"></param>
        public void FindFile(WinQuestionEntity winquestion)
        {
            //将正确答案，分值取出来，传给studentRecord

          
            winquestion.questionFlag = "查找文件";

            DataTable winQuestionDt = winquestionbll.LoadWindowsByFlag(winquestion);

           WinQuestionRecordEntity studentRecord = new WinQuestionRecordEntity();

            studentRecord.studentID = FrmLogin.studentID;
            string correctAnswer;
            string fraction;
            string examAnswer;
            //循环遍历正确答案
            for (int i = 0; i < winQuestionDt.Rows.Count; i++)
            {
                //将考生ID传到studentRecord实体
                studentRecord.studentID = FrmLogin.studentID;
                //将试题的ID取出来
                studentRecord.questionID = Convert.ToDouble(winQuestionDt.Rows[i]["questionID"]);
                //将题的分数取出来
                fraction = winQuestionDt.Rows[i]["fraction"].ToString();
                //将正确答案取出来
                correctAnswer = winQuestionDt.Rows[i]["correctAnswer"].ToString();

                string str = @"D:\计算机一级考生文件\winkt" + correctAnswer;
                if (File.Exists(str))
                {
                    //加分
                    studentRecord.fraction = Convert.ToDouble(fraction);
                    examAnswer = correctAnswer;
                    studentRecord.examAnswer = examAnswer;
                }
                else
                {
                    //不加分
                    studentRecord.fraction = 0;
                    studentRecord.examAnswer = "";
                }
                winquestionbll.ReturnScore(studentRecord);
            }
        }
        #endregion
    }
}
