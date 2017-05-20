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
using System.Threading;
using System.Text.RegularExpressions;

namespace NCRE学生考试端V1._0
{
    public class WindowsFindStartFiles
    {
        private WinQuestionEntityBLL winquestionbll = new WinQuestionEntityBLL();

        #region"开头查找文件-韩梦甜-2015-11-23"
        /// <summary>
        /// 开头，后缀名查找文件
        /// </summary>
        /// <param name="winquestion"></param>
        public void FindStartFiles(WinQuestionEntity winquestion)
        {
            //将正确答案，分值取出来，传给studentRecord

            winquestion.questionFlag = "开头查找文件";

            DataTable winQuestionDt = winquestionbll.LoadWindowsByFlag(winquestion);

            WinQuestionRecordEntity studentRecord = new WinQuestionRecordEntity();
            studentRecord.studentID = FrmLogin.studentID;

            string correctAnswer;
            string fraction;
            string examAnswer = "";
            string questionContent;

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

                questionContent = winQuestionDt.Rows[i]["questionContent"].ToString();


                Regex re = new Regex("(?<=“).*?(?=”)", RegexOptions.None);
                MatchCollection mc = re.Matches(questionContent);
                string str = @"D:\计算机一级考生文件\winkt\" + mc[1].ToString();
                //容灾处理，如果查不到路径，则为0分-韩梦甜-2014-12-6
                if (Directory.Exists(str) == false)
                {
                    studentRecord.fraction = 0;
                    studentRecord.examAnswer = "";
                }
                else
                {
                    var files = Directory.GetFiles(str, mc[0].ToString() + "*");

                    //获取考生答案
                    foreach (var file in files)
                        examAnswer += file;
                    studentRecord.examAnswer = examAnswer;

                    //判断答案是否正确
                    if (examAnswer == correctAnswer)
                    {
                        //加分
                        studentRecord.fraction = Convert.ToDouble(fraction);

                    }
                    else
                    {
                        //不加分
                        studentRecord.fraction = 0;

                    }
                }
                winquestionbll.ReturnScore(studentRecord);
            }
        }
        #endregion
    }
}
