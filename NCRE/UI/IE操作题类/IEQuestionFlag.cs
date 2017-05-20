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
    public class IEQuestionFlag
    {
        private IEQuestionEntityBLL iequestionbll = new IEQuestionEntityBLL();
        private IEQuestionEntity iequestion=new IEQuestionEntity ();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();

        #region"根据题的类型，调用相应的判分-韩梦甜-2015-11-20"
        /// <summary>
        /// 根据题的类型，调用相应的判分
        /// </summary>
        /// <param name="iequestion"></param>
        public void SwitchQuestionFlag(IEQuestionEntity iequestion)
        {
            DataTable dt = new DataTable();
       
            dt = iequestionbll.LoadIEQuestion(iequestion );
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                String collegeId = MyInfo.MycollegeID();
                if (examconfigbll.IsTableExist("IEQuestionRecordEntity_" + collegeId) == false)
                {
                    break;
                }
                else
                {
                    switch (questionflag)
                    {

                        case "查看收藏夹":
                            IEFindFavorites iefindfavorites = new IEFindFavorites();
                            iefindfavorites.FindFavorites(iequestion);
                            break;

                        case "起始页":
                            IEFindStartPage iefindstartpage = new IEFindStartPage();
                            iefindstartpage.FindStartPage(iequestion);
                            break;

                        case "查找文件":
                            IEFindFiles iefindfiles = new IEFindFiles();
                            iefindfiles.FindFile(iequestion);
                            break;

                        case "查看历史记录":
                            IEFindHistoryRecord iefindhistoryrecord = new IEFindHistoryRecord();
                            iefindhistoryrecord.FindHistoryRecord(iequestion);
                            break;                    

                    }
                }
            }
        }
        #endregion
    }
}
