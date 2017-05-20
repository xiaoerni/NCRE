using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using Model;
using System.Data;
using System.IO;


namespace NCRE学生考试端V1._0
{
    public static class MyInfo
    {


        #region  获取学生的ID为全局的方法--周洲==2015年11月21日

        /// <summary>
        /// 获取学生的ID为全局的方法--周洲==2015年11月21日
        /// </summary>
        /// <returns></returns>
        public static string MystudentID()
        {
            return FrmLogin.studentID ;
        }
        #endregion

        #region    //获取全局变量CollegeID-----周洲---2015年11月21日
        /// <summary>
        ///    //获取全局变量CollegeID-----周洲---2015年11月21日
        /// </summary>
        /// <returns></returns>
        public static string MycollegeID()
        {

            StudentInfoEntityBLL studentinfobll = new StudentInfoEntityBLL();
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            //获取学生信息
            studentinfo.studentID = MyInfo.MystudentID();
            DataTable dt = studentinfobll.SelectStudentInfoByID(studentinfo);
            //获取全局变量CollegeID
            return dt.Rows[0]["collegeID"].ToString();
        }
        #endregion

        #region 获取全局的学生试卷类型--周洲--2015年11月21日


        /// <summary>
        /// 获取全局的学生试卷类型--周洲--2015年11月21日
        /// </summary>
        /// <returns></returns>
        public static string MyPaperType()
        {
            StudentInfoEntityBLL studentinfobll = new StudentInfoEntityBLL();
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = MyInfo.MystudentID();
            //获取学生试卷类型信息
            DataTable papertypedt = studentinfobll.SelectPaperTypebyStudentID(studentinfo);
            //给全局变量赋值，传递试卷类型信息
            return papertypedt.Rows[0]["PaperType"].ToString();
        }
        #endregion


        #region 判断该文档是否存在于此路径下 2015年12月11日15:19:47 李少然
        /// <summary>
        /// //判断该文档是否存在于此路径下
        /// </summary>
        /// <param name="path">文档路径</param>
        /// <returns>true</returns>
        public static Boolean exitsDoc(string path)
        {
            if (File.Exists(path))
            {
                return true;
            }
            else
            {
                return false;
            }
        } 
        #endregion
    }
}
