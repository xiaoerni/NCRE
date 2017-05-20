/********************************************************************************** 
     * 开发人:邱慕夏
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/23 14:19:11 
     *开发版本：V1.0
 **********************************************************************************/

using Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace DAL
{
    public class StudentBindPaperTypeDAL
    {
        private SQLHelper sqlhelper = null;
        public StudentBindPaperTypeDAL()
        {
            sqlhelper = new SQLHelper();
        }


        #region 根据学生的ID抽提的时候进行查询的sql语句---邱慕夏--2015年11月22日
        /// <summary>
        /// 根据学生的ID抽提的时候进行查询的sql语句---邱慕夏--2015年11月22日
        /// </summary>
        /// <param name="studentinfo">学生的ID</param>
        /// <returns></returns>
        public DataTable SelectAllMajor(String studentID)
        {
            DataTable dt = new DataTable();
            string sql = "select PaperType,CollegeID from StudentBindPaperTypeEntity where StudentID=@StudentID and IsUse=@IsUse";
            SqlParameter[] paras = new SqlParameter[] {
                new SqlParameter ("@StudentID",studentID),
                new SqlParameter ("@IsUse",'1')};
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region 根据学生的ID抽提的时候进行查询的sql语句---邱慕夏--2015年11月22日
        /// <summary>
        /// insert--邱慕夏
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns>是否添加成功</returns>
        public int InsertRecord(StudentBindPaperTypeEntity studentBindPaperType)
        {
            string sql = "Insert into StudentBindPaperTypeEntity(StudentID,PaperType,IsUse,CollegeID) values(@StudentID,@PaperType,@IsUse,@CollegeID)";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentBindPaperType.StudentID),
                new SqlParameter ("@PaperType",studentBindPaperType.PaperType),
                new SqlParameter ("@IsUse",studentBindPaperType.IsUse),
                new SqlParameter ("@CollegeID",studentBindPaperType.CollegeID)
            };
            int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            return flag;
        }
        #endregion


        #region 根据学生的ID去重---邱慕夏--2015年11月22日
        /// <summary>
        /// select--邱慕夏
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns>是否已经存在</returns>
        public Boolean SelectRecord(StudentBindPaperTypeEntity studentBindPaperType)
        {
            string sql = "Select * from StudentBindPaperTypeEntity where StudentID=@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentBindPaperType.StudentID),
            };
            DataTable dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            if (dt.Rows.Count == 0)
            {
                return true;
            }
            else {
                return false;
            }
        }
        #endregion

        #region 批量根据学生的ID抽提的时候进行查询的sql语句---邱慕夏--2015年11月22日
        /// <summary>
        /// insert--邱慕夏
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns>是否添加成功</returns>
        public int InsertRecordList(List<StudentBindPaperTypeEntity> studentBindPaperTypeList)
        {
            int flag = 0;
            for (int i = 0; i < studentBindPaperTypeList.Count; i++)
            {

                string sql = "Insert into StudentBindPaperTypeEntity(StudentID,PaperType,IsUse,CollegeID) values(@StudentID,@PaperType,@IsUse,@CollegeID)";
                SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentBindPaperTypeList[i].StudentID),
                new SqlParameter ("@PaperType",studentBindPaperTypeList[i].PaperType),
                new SqlParameter ("@IsUse",studentBindPaperTypeList[i].IsUse),
                new SqlParameter ("@CollegeID",studentBindPaperTypeList[i].CollegeID)
                };
                flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            }
            return flag;
        }
        #endregion
    }
}
