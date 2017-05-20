using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using BLL;
using Model;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Win32;
using SHDocVw;


namespace NCRE学生考试端V1._0
{
    public class IEFindStartPage
    {
        private IEQuestionEntityBLL iequestionbll = new IEQuestionEntityBLL();

        #region"起始页-韩梦甜-2015-11-20"
        /// <summary>
        /// 起始页
        /// </summary>
        /// <param name="iequestion"></param>
        public void FindStartPage(IEQuestionEntity iequestion)
        {
            //将正确答案，分值取出来，传给studentRecord
            
            iequestion.questionFlag = "起始页";

            DataTable ieQuestionDt = iequestionbll.LoadIEByFlag (iequestion);

            IEQuestionRecordEntity studentRecord = new IEQuestionRecordEntity();

            studentRecord.studentID = FrmLogin.studentID;

            string correctAnswer;
            string fraction;
            string examAnswer;


            //循环遍历正确答案
            for (int i = 0; i < ieQuestionDt.Rows.Count; i++)
            {
                //将考生ID传到studentRecord实体
                studentRecord.studentID = FrmLogin.studentID;
                //将试题的ID取出来
                studentRecord.questionID = Convert.ToDouble(ieQuestionDt.Rows[i]["questionID"]);
                //将题的分数取出来
                fraction = ieQuestionDt.Rows[i]["fraction"].ToString();
                //将正确答案取出来
                correctAnswer = ieQuestionDt.Rows[i]["correctAnswer"].ToString();

                ShellWindows shellwindows = new ShellWindows();
                foreach (InternetExplorer ie in shellwindows)
                {
                    string filename = Path.GetFileNameWithoutExtension(ie.FullName).ToLower();
                    if (filename.Equals("iexplore"))
                    {
                        string str = ie.LocationURL.ToString();

                        if (str == correctAnswer)
                        {
                            //加分
                            studentRecord.fraction= Convert.ToDouble(fraction);
                            examAnswer = correctAnswer;
                            studentRecord.examAnswer = examAnswer;
                        }
                        else
                        {
                            //不加分
                            studentRecord.fraction = 0;
                            studentRecord.examAnswer =str;
                        }
                        iequestionbll.ReturnScore(studentRecord);
                    }
                }
            }
        }
        #endregion
    }
}
                  

