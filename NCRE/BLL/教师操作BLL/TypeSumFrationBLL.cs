//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using Model;
//using System.Data;
//using DAL;

//namespace BLL
//{
//    public class TypeSumFrationBLL
//    {
//        #region
//        /// <summary>
//        /// word——李芬
//        /// </summary>
//        /// <param name="studentinfo"></param>
//        /// <returns></returns>
//        public DataTable WordfSumFration(StudentInfoEntity studentinfo)
//        {
//            TypesumfrationDAL typesumfrationdal = new TypesumfrationDAL();
//            return typesumfrationdal.wordsumfrationdal(studentinfo);

//        }
//        #endregion
//        #region
//        /// <summary>
//        /// win——李芬
//        /// </summary>
//        /// <param name="studentinfo"></param>
//        /// <returns></returns>
//        public DataTable winfSumFration(StudentInfoEntity studentinfo)
//        {
//            TypesumfrationDAL winesumfrationdal = new TypesumfrationDAL();
//            return winesumfrationdal.winsumfrationdal(studentinfo);

//        }
//        #endregion
//    }
//}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;
using System.Data;
using DAL;

namespace BLL
{
    /// <summary>
    /// 根据学号计算所有题型总分-王荣晓、李芬-2015年11月16日
    /// </summary>
    public class TypeSumFrationBLL
    {


        public DataTable studentIDscore(StudentInfoEntity studentinfo)
        {
            TypesumfrationDAL tyscore = new TypesumfrationDAL();
            return tyscore.selectstudentID(studentinfo);


        }
        #region   调用D层计算word题型总分
        /// <summary>
        /// 计算word题型总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable WordSumFration(StudentInfoEntity studentinfo)
        {
            TypesumfrationDAL typesumfrationdal = new TypesumfrationDAL();
            return typesumfrationdal.wordsumfrationdal(studentinfo);

        }
        #endregion

        #region 选择题成绩汇总

        public DataTable SelectsumFration(StudentInfoEntity studentinfo)
        {
            TypesumfrationDAL typesumfrationdal = new TypesumfrationDAL();
            return typesumfrationdal.Selectsumfration(studentinfo);
        }
        #endregion

        #region  计算PPT题型总分
        /// <summary>
        /// 计算PPT题型总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable PPTSumFration(StudentInfoEntity studentinfo)
        {
            TypesumfrationDAL typesumfrationdal = new TypesumfrationDAL();
            return typesumfrationdal.pptsumfrationdal(studentinfo);

        }

        #endregion

        #region  计算IE题型总分
        /// <summary>
        /// 计算IE题型总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable IESumFration(StudentInfoEntity studentinfo)
        {
            TypesumfrationDAL typesumfrationdal = new TypesumfrationDAL();
            return typesumfrationdal.IEsumfrationdal(studentinfo);

        }

        #endregion

        #region  计算windows题型总分
        /// <summary>
        /// 计算windows题型总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable windowsSumFration(StudentInfoEntity studentinfo)
        {
            TypesumfrationDAL typesumfrationdal = new TypesumfrationDAL();
            return typesumfrationdal.winsumfrationdal(studentinfo);

        }

        #endregion

        #region  计算Excel题型总分
        /// <summary>
        /// 计算Excel题型总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable ExcelSumFration(StudentInfoEntity studentinfo)
        {
            TypesumfrationDAL typesumfrationdal = new TypesumfrationDAL();
            return typesumfrationdal.excelsumfrationdal(studentinfo);

        }

        #endregion

        #region 汇总每位同学的总成绩--李芬
        public DataTable SumFration(StudentInfoEntity studentinfo)
        {
            TypesumfrationDAL typesumfrationdal = new TypesumfrationDAL();
            return typesumfrationdal.sumfrationdal(studentinfo);

        }
        #endregion

    }
}
