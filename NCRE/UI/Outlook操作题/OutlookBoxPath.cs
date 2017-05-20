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
    public class OutlookBoxPath
    {
        private IEQuestionEntityBLL iequestionbll = new IEQuestionEntityBLL();

        #region"查看邮件附件-韩梦甜-2015-11-21"
        /// <summary>
        /// 查看邮件附件
        /// </summary>
        /// <param name="iequestion"></param>
        public void FindBoxPath(IEQuestionEntity iequestion)
        {
            iequestion.questionFlag = "邮件附件";
            DataTable ieQuestionDt = iequestionbll.LoadIEByFlag(iequestion);

            IEQuestionRecordEntity studentRecord = new IEQuestionRecordEntity();

            studentRecord.studentID = FrmLogin.studentID;
            string fraction;

            //循环遍历正确答案
            for (int i = 0; i < ieQuestionDt.Rows.Count; i++)
            {
                //将考生ID传到studentRecord实体
                studentRecord.studentID = FrmLogin.studentID;
                //将试题的ID取出来
                studentRecord.questionID = Convert.ToDouble(ieQuestionDt.Rows[i]["questionID"]);
                //将题的分数取出来
                fraction = ieQuestionDt.Rows[i]["fraction"].ToString();
                //将考生答案保存
                studentRecord.examAnswer = frmNewM.BoxPath;
                if (frmNewM.BoxPath == ieQuestionDt.Rows[i]["correctAnswer"].ToString())
                {
                    studentRecord.fraction = Convert.ToDouble(fraction);

                }
                else
                {
                    studentRecord.fraction = 0;
                }
                iequestionbll.ReturnScore(studentRecord);


            }
        }
        #endregion
    }
}
