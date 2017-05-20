using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using BLL;
using Model;
using System.Reflection;

//此类负责加载题的信息：题目：正确答案，分值

namespace NCRE学生考试端V1._0
{
   public  class WindowsLoadInfo
    {
       public WinQuestionEntityBLL winquestionbll = new WinQuestionEntityBLL();

       #region 根据试题类型获取题库中的试题--周洲--2015年11月21日
       /// <summary>
       /// 获取考试内容界面
       /// </summary>
       /// <param name="winQuestion">传递考试类型</param>
       /// <returns></returns>
       public DataTable LoadQuestionContent(WinQuestionEntity wininfo)
       {
           return winquestionbll.LoadWindowsQuestion(wininfo);
       }
       #endregion

       
    }
}
