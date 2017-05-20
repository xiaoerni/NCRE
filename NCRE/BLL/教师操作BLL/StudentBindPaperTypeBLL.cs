/********************************************************************************** 
     * 开发人:邱慕夏
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/23 14:12:19 
     *开发版本：V1.0
 **********************************************************************************/

using Model;
using DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace BLL
{
    public class StudentBindPaperTypeBLL
    {
        private StudentBindPaperTypeDAL studentBindPaperTypeDAL;
        public StudentBindPaperTypeBLL()
        {
            studentBindPaperTypeDAL = new StudentBindPaperTypeDAL();
        }

        #region 根据学生的ID抽提的时候进行查询的sql语句---邱慕夏--2015年11月22日
        /// <summary>
        /// 根据学生的ID抽提的时候进行查询的sql语句---邱慕夏--2015年11月22日
        /// </summary>
        /// <param name="studentinfo">学生的ID</param>
        /// <returns></returns>
        public DataTable SelectAllMajor(String studentID)
        {
            return studentBindPaperTypeDAL.SelectAllMajor(studentID);
        }
        #endregion


        #region insert---邱慕夏--2015年11月22日
        /// <summary>
        /// insert--邱慕夏
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns>是否添加成功</returns>
        public int InsertRecord(StudentBindPaperTypeEntity studentBindPaperType)
        {
            return studentBindPaperTypeDAL.InsertRecord(studentBindPaperType);
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
            return studentBindPaperTypeDAL.SelectRecord(studentBindPaperType);
        }
        #endregion

        #region 批量insert---邱慕夏--2015年11月22日
        /// <summary>
        /// insert--邱慕夏
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns>是否添加成功</returns>
        public int InsertRecordList(List<StudentBindPaperTypeEntity> studentBindPaperTypeList)
        {
            return studentBindPaperTypeDAL.InsertRecordList(studentBindPaperTypeList);
        }
        #endregion
    }
}
