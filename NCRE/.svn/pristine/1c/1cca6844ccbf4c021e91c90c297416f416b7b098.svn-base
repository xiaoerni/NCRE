using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using BLL;
using Model;

namespace NCRE学生考试端V1._0
{
    public class IELoadInfo
    {
        public IEQuestionEntityBLL iequestionbll = new IEQuestionEntityBLL();

        #region 根据题型选择出IE的题干--周洲--2015年11月21日
        /// <summary>
        ///根据题型选择出IE的题干--周洲--2015年11月21日
        /// </summary>
        /// <param name="winQuestion">传递考试类型</param>
        /// <returns></returns>
        public DataTable LoadQuestionContent(IEQuestionEntity ieinfo)
        {
            return iequestionbll.LoadIEQuestion(ieinfo);
        }
        #endregion
    }
}
