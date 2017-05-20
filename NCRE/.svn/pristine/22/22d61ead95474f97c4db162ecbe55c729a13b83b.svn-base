using DAL;
using Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BLL
{
    /// <summary>
    /// 主要存放  学生相关的信息
    /// </summary>
    public class StudentInfoBLL
    {
        StudentInfoEntityDAL studentDal = new StudentInfoEntityDAL();

        /// <summary>
        /// 通过学号查询  学生信息
        /// </summary>
        /// <param name="studentId"></param>
        /// <returns></returns>
        public StudentInfoEntity GetStudentById(string studentId) {
            return studentDal.GetStudentById(studentId);
        }
    }
}
