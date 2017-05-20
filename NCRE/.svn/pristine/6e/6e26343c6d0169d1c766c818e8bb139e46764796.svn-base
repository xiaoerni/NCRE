using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using Model;
using System.Data;
using System.Data.SqlClient;

namespace NCRE学生考试端V1._0
{
    public class IEFindFavorites
    {
        private IEQuestionEntityBLL iequestionbll = new IEQuestionEntityBLL();

        #region"查看收藏夹-韩梦甜-2015-11-20"
        /// <summary>
        /// 查看收藏夹
        /// </summary>
        /// <param name="iequestion"></param>
        public void FindFavorites(IEQuestionEntity iequestion)
        {
            //将正确答案，分值取出来，传给studentRecord
          
            iequestion.questionFlag = "查看收藏夹";

            DataTable ieQuestionDt = iequestionbll.LoadIEByFlag(iequestion);

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

                string favorfolder = Environment.GetFolderPath(Environment.SpecialFolder.Favorites);
                //System.Diagnostics.Process.Start("explorer.exe", favorfolder);
                string str = favorfolder + @"\" + correctAnswer;
                //查看是否有该网页
                if (System.IO.File.Exists(str))
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
                iequestionbll.ReturnScore(studentRecord);
            }
        }
        #endregion
    }
}
      