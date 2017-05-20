using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.IO;
using IWshRuntimeLibrary;
using System.Threading;

namespace NCRE学生考试端V1._0
{
    public class WindowsFindIWshShortcut
    {
        private WinQuestionEntityBLL winquestionbll = new WinQuestionEntityBLL();

        #region"查找快捷方式-韩梦甜-2015-11-20"
        /// <summary>
        /// 查找快捷方式
        /// </summary>
        /// <param name="winquestion"></param>
        public void FindIWshShortcut(WinQuestionEntity winquestion)
        {
            //将正确答案，分值取出来，传给studentRecord
          
            winquestion.questionFlag = "查找快捷方式";

            DataTable winQuestionDt = winquestionbll.LoadWindowsByFlag(winquestion);

           WinQuestionRecordEntity studentRecord = new WinQuestionRecordEntity();

            studentRecord.studentID = FrmLogin.studentID;
            string correctAnswer;
            string fraction;
            string examAnswer;
            WshShell shell = new WshShell();

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
                //查看快捷方式是否存在
                IWshShortcut shourtcut = (IWshShortcut)shell.CreateShortcut(str);

                if (System.IO.File.Exists(shourtcut.TargetPath))
                {
                    //加分
                    studentRecord.fraction = Convert.ToDouble(fraction);
                    examAnswer = correctAnswer;
                    studentRecord.examAnswer =examAnswer ;
                }
                else
                {
                    //不加分
                    studentRecord.fraction= 0;
                    studentRecord.examAnswer = "";
                }
                winquestionbll.ReturnScore(studentRecord);
            }
        }
        #endregion

    }
}
