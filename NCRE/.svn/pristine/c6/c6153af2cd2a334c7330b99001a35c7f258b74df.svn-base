using Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace DAL
{
    /// <summary>
    /// 学院信息
    /// </summary>
    public class CollegeDAL
    {
        private SQLHelper sqlhelper = null;

        public CollegeDAL()
        {
            sqlhelper =new SQLHelper(); 
        }

        TBToList<CollegeEntity> dtToCollegeList = new TBToList<CollegeEntity>();
        TBToList<StudentInfoEntity> dtToStudentList = new TBToList<StudentInfoEntity>();

        /// <summary>
        /// 读取全部单选题
        /// </summary>
        /// <returns>单选题集合List</returns>
        public List<CollegeEntity> GetAllCollege()
        {
            //1,查询所有的选择题
            DataTable dt = new DataTable();
            string sql = "select * from CollegeEntity";

            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);

            //2,把查询结果组织成  实体List
            List<CollegeEntity> lsSelectQuestionEntity = new List<CollegeEntity>();
            lsSelectQuestionEntity = dtToCollegeList.ToList(dt);

            return lsSelectQuestionEntity;
        }

        /// <summary>
        /// 根据学院ID获取 该学院的学生列表
        /// </summary>
        /// <param name="pCollege"></param>
        /// <returns></returns>
        public List<StudentInfoEntity> GetStudentByCollege(CollegeEntity pCollege)
        {
            //1,根据学院ID获取 该学院的学生列表
            DataTable dt = new DataTable();
            string sql = "select * from StudentInfoEntity where collegeID=@collegeID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter ("@collegeID",pCollege.collegeID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);

            //2,把查询结果组织成  实体List
            List<StudentInfoEntity> lsSelectQuestionRecordEntity = new List<StudentInfoEntity>();
            lsSelectQuestionRecordEntity = dtToStudentList.ToList(dt);

            return lsSelectQuestionRecordEntity;
        }
    }
}
