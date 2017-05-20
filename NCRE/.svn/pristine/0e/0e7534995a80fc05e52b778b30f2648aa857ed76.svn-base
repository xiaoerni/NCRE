using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DAL;
using Model;
using System.Data;
using System.Data.SqlClient;


namespace BLL
{
    public class StudentInfoEntityBLL
    {
            private StudentInfoEntityDAL  studentinfodal = null;
            public StudentInfoEntityBLL()
            {
                studentinfodal = new StudentInfoEntityDAL();
            }

            #region 选择学生信息通过学号---周洲--2015-11-21


            /// <summary>
            /// 选择学生信息通过学号---周洲--2015-11-21
            /// </summary>
            /// <param name="studentinfo"></param>
            /// <returns></returns>
            public DataTable SelectStudentInfoByID(StudentInfoEntity studentinfo)
            {
                return studentinfodal.SelectStudentInfoByID(studentinfo);
            }
            #endregion

            #region 通过studentID选择学生的试卷类型（A,B,C）--周洲--2015年11月21日
            /// <summary>
            /// 通过studentID选择学生的试卷类型（A,B,C）--周洲--2015年11月21日
            /// </summary>
            /// <param name="studentinfo"></param>
            /// <returns></returns>
            public DataTable SelectPaperTypebyStudentID(StudentInfoEntity studentinfo)
            {
                return studentinfodal.SelectPaperTypebyStudentID(studentinfo);
            }
            #endregion

    }
}
