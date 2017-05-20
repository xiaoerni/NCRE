using Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace DAL
{
    public class SelectQuestionEntityDAL
    {
        private SQLHelper sqlhelper = null;

        public SelectQuestionEntityDAL()
        {
            sqlhelper = new SQLHelper();
        }

        TBToList<SelectQuestionEntity> dtToList = new TBToList<SelectQuestionEntity>();

        /// <summary>
        /// 读取全部单选题
        /// </summary>
        /// <returns>单选题集合List</returns>
        public List<SelectQuestionEntity> LoadAllSelectQuestion()
        {
            //1,查询所有的选择题
            DataTable dt = new DataTable();
            string sql = "select * from SelectQuestionEntity";

            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);

            //2,把查询结果组织成  实体List
            List<SelectQuestionEntity> lsSelectQuestionEntity = new List<SelectQuestionEntity>();
            lsSelectQuestionEntity = dtToList.ToList(dt);

            return lsSelectQuestionEntity;
        }

        //TODO 在线编辑 写入选择题题库

    }
}
