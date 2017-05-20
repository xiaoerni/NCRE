using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using BLL;
using Model;
using System.Reflection;

namespace NCRE学生考试端V1._0
{
    public class PptLoadinfo
    {
        public PptQuestionEntityBLL pptquestionbll = new PptQuestionEntityBLL();



        /// <summary>
        /// 获取考试内容的界面
        /// </summary>
        /// <param name="pptinfo">传递考试类型</param>
        /// <returns>返回考试内容</returns>
        public DataTable LoadQuestionContent(PptQuestionEntity pptinfo)
        {
            return pptquestionbll.LoadPptQuestion(pptinfo);
        }
    }
}
