using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using BLL;
using Model;
using System.Reflection;
//这个类负责加载题的信息：题目，正确答案，分值

namespace NCRE学生考试端V1._0
{
    public class WordLoadinfo 
    {
        public WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        #region 利用全局变量，从题库中加载word试题--周洲--2015年11月21日
        /// <summary>
        /// 获取考试内容的界面
        /// </summary>
        /// <param name="wordinfo">传递考试类型</param>
        /// <returns>返回考试内容</returns>
        public DataTable LoadQuestionContent(WordQuestionEntity wordinfo)
        {
            //根据学号查询该学生要考的试题和试卷类型，
            return wordquestionbll.LoadWordQuestion(wordinfo);
        }
        #endregion
    }
}
