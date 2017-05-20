using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Model;
using BLL;

namespace NCRE学生考试端V1._0.Windows操作题类
{
   public  class WindowsReturnPaper
   {

       #region"返回判分之后的信息-韩梦甜-2015-11-20"
       
       /// <summary>
       /// 返回判分之后的信息
       /// </summary>
       /// <param name="studentRecord">得分，学生ID，考试答案，学生答案</param>
       /// <returns></returns>
       public int ReturnScore(WinQuestionRecordEntity studentRecord)
       {
           WinQuestionEntity winQuestion = new WinQuestionEntity();
     
           WinQuestionEntityBLL winquestionbll = new WinQuestionEntityBLL();
           int flag = winquestionbll.ReturnScore(studentRecord);
           return flag;
       }
       #endregion
   }
}
