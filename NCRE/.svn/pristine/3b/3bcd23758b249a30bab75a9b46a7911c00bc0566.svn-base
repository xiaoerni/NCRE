using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Model;
using BLL;

namespace NCRE学生考试端V1._0.IE操作题类
{
    public class IEReturnPaper
    {
        #region"返回判分之后的信息-韩梦甜-2015-11-20"

        /// <summary>
        /// 返回判分之后的信息
        /// </summary>
        /// <param name="studentRecord">得分，学生ID，考试答案，学生答案</param>
        /// <returns></returns>
        public int ReturnScore(IEQuestionRecordEntity studentRecord)
        {
            //实例化一个判分类，获取返回的得分和考生答案
            IEQuestionEntity winQuestion = new IEQuestionEntity();
           
            IEQuestionEntityBLL iequestionbll = new IEQuestionEntityBLL();
            int flag = iequestionbll.ReturnScore(studentRecord);
            return flag;
        }
        #endregion
    }
}
