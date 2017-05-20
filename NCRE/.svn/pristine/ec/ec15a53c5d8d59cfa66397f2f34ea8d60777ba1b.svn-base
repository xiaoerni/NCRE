using Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using DAL;

namespace DAL
{
    public class TypesumfrationDAL
    {
        #region 实例化一个sqlhelper
        private SQLHelper sqlhelper = null;

        WordQuestionEntityDAL wordtypesum = new WordQuestionEntityDAL();
        /// <summary>
        /// 实例化一个sqlhelper
        /// </summary>
        public TypesumfrationDAL()
        {
            sqlhelper = new SQLHelper();

        }
        #endregion

        #region word类汇总总分-李芬

        /// <summary>
        /// word类汇总总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable wordsumfrationdal(StudentInfoEntity studentinfo)
        {
            String which = selectCollegeID(studentinfo).ToString();
            //查询学生的试卷类型/类型总分值
            SqlParameter[] paras = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string sumText = "select sum(convert(float,Fration)) from WordQuestionRecordEntity_" + which + " where StudentID=@studentID group by StudentID";

            DataTable wordsumFrationDt;
            wordsumFrationDt = sqlhelper.ExecuteQuery(sumText, paras, CommandType.Text);
            string sumFration;
            //if (true)
            //{

            //}
            sumFration = wordsumFrationDt.Rows[0][0].ToString();


            //查询PaperType
            SqlParameter[] parastype = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string papertypeText = "select PaperType from WordQuestionRecordEntity_" + which + " where StudentID=@studentID ";

            DataTable wordpapertypeDT;
            wordpapertypeDT = sqlhelper.ExecuteQuery(papertypeText, parastype, CommandType.Text);
            string paperType;
            paperType = wordpapertypeDT.Rows[0][0].ToString();



            //查询学生QuestionTypeID
            //SqlParameter[] QuestionType = new SqlParameter[] { new SqlParameter("@paperType", paperType) };
            //string QuestionTypeIDText = "select QuestionTypeID,PaperType from WordQuestionRecordEntity_" + which + " where PaperType=@paperType";
            //DataTable wordQuestionTypeIDDt;
            //wordQuestionTypeIDDt = sqlhelper.ExecuteQuery(QuestionTypeIDText, QuestionType, CommandType.Text);
            //string QuestionTypeID;
            //QuestionTypeID = wordQuestionTypeIDDt.Rows[0]["QuestionTypeID"].ToString();


            //在WordSumFration表中插入数据
            DateTime dtime = System.DateTime.Now;

            SqlParameter[] parasinsertword = new SqlParameter[]{ new SqlParameter("@StudentID",studentinfo.studentID),new SqlParameter("@PaperType",paperType),new SqlParameter("@QuestionTypeID","NULL"),new SqlParameter("@sumfartion",sumFration),new SqlParameter("@TimeStamp",dtime)
            };
            string insertword = "insert into WordSumFration values(@StudentID,@PaperType,@QuestionTypeID,@sumfartion,@TimeStamp)";
            int numword = sqlhelper.ExecuteNonQuery(insertword, parasinsertword, CommandType.Text);
            if (numword > 0)
            {
                return wordsumFrationDt;
            }
            return null;
        }
        #endregion
        #region excel类汇总总分

        /// <summary>
        /// excel类汇总总分-李芬
        /// 
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable excelsumfrationdal(StudentInfoEntity studentinfo)
        {
            String which = selectCollegeID(studentinfo).ToString();

            //查询学生的试卷类型/类型总分值
            SqlParameter[] paras = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string sumText = "select  sum(convert(float,Fration))  from ExcelQuestionRecordEntity_" + which + " where StudentID=@studentID group by StudentID";

            DataTable excelsumFrationDt;
            excelsumFrationDt = sqlhelper.ExecuteQuery(sumText, paras, CommandType.Text);
            string sumFration;
            sumFration = excelsumFrationDt.Rows[0][0].ToString();

            //SqlParameter[] excelparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            //string excelpapertypeText = "select PaperType from ExcelQuestionRecordEntity _" + which + " where StudentID=@studentID";

            //DataTable excelpapertypeDt;
            //excelpapertypeDt = sqlhelper.ExecuteQuery(excelpapertypeText, excelparas, CommandType.Text);

            //string PaperType;
            //PaperType = excelpapertypeDt.Rows[0]["PaperType"].ToString();


            //查询学生QuestionTypeID
            //SqlParameter[] excelquestiontypeparas = new SqlParameter[] { new SqlParameter("@PaperType", PaperType) };
            //string QuestionTypeIDText = "select QuestionTypeID from ExcelQuestionEntity_" + which + " where PaperType=@PaperType";
            //DataTable excelQuestionTypeIDDt;
            //excelQuestionTypeIDDt = sqlhelper.ExecuteQuery(QuestionTypeIDText, excelquestiontypeparas, CommandType.Text);
            //string QuestionTypeID;
            //QuestionTypeID = excelQuestionTypeIDDt.Rows[0]["QuestionTypeID"].ToString();

            //在ExcelSumFration表中插入数据
            DateTime dtime = System.DateTime.Now;
            SqlParameter[] parasinsertexcel = new SqlParameter[]{ new SqlParameter("@StudentID",studentinfo.studentID),new SqlParameter("@PaperType","NULL"),new SqlParameter("@QuestionTypeID","NULL"),new SqlParameter("@sumfartion",sumFration),new SqlParameter("@TimeStamp",dtime)
            };
            string insertexcel = "insert into ExcelSumFration values(@StudentID,@PaperType,@QuestionTypeID,@sumfartion,@TimeStamp)";
            int numexcel = sqlhelper.ExecuteNonQuery(insertexcel, parasinsertexcel, CommandType.Text);

            return excelsumFrationDt;

        }
        #endregion
        #region win类汇总总分-李芬

        /// <summary>
        /// win类汇总总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable winsumfrationdal(StudentInfoEntity studentinfo)
        {
            String which = selectCollegeID(studentinfo).ToString();
            //查询学生的类型总分值
            SqlParameter[] paras = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string sumText = "select  sum(convert(float,fraction)) from WinQuestionRecordEntity_" + which + " where studentID=@studentID group by studentID";
            DataTable winsumFrationDt;
            winsumFrationDt = sqlhelper.ExecuteQuery(sumText, paras, CommandType.Text);
            string sumFration;
            sumFration = winsumFrationDt.Rows[0][0].ToString();




            //在winSumFration表中插入数据
            DateTime dtime = System.DateTime.Now;
            SqlParameter[] parasinsertwin = new SqlParameter[]{ new SqlParameter("@StudentID",studentinfo.studentID),new SqlParameter("@PaperType","NULL"),new SqlParameter("@QuestionTypeID","NULL"),new SqlParameter("@sumfartion",sumFration),new SqlParameter("@TimeStamp",dtime)
            };
            string insertwin = "insert into WinSumFration values(@StudentID,@PaperType,@QuestionTypeID,@sumfartion,@TimeStamp)";
            int numwin = sqlhelper.ExecuteNonQuery(insertwin, parasinsertwin, CommandType.Text);
            return winsumFrationDt;
        }
        #endregion
        #region ppt类汇总总分-李芬

        /// <summary>
        /// ppt类汇总总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable pptsumfrationdal(StudentInfoEntity studentinfo)
        {
            String which = selectCollegeID(studentinfo).ToString();
            //查询学生的试卷类型/类型总分值
            SqlParameter[] paras = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string sumText = "select sum(convert(float,Fration)) from PptQuestionRecordEntity_" + which + " where StudentID=@studentID group by StudentID";

            DataTable pptsumFrationDt;
            pptsumFrationDt = sqlhelper.ExecuteQuery(sumText, paras, CommandType.Text);
            string sumFration;
            sumFration = pptsumFrationDt.Rows[0][0].ToString();



            //SqlParameter[] pptpapertypeparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            //string pptpapertypeText = "select PaperType from PptQuestionRecordEntity_" + which + "    where StudentID=@studentID ";

            //DataTable pptpapertypeDt;
            //pptpapertypeDt = sqlhelper.ExecuteQuery(pptpapertypeText, pptpapertypeparas, CommandType.Text);

            //string pptPaperType;
            //pptPaperType = pptpapertypeDt.Rows[0]["PaperType"].ToString();


            ////查询学生QuestionTypeID
            //SqlParameter[] pptpapertype = new SqlParameter[] { new SqlParameter("@papertype", pptPaperType) };
            //string pptQuestionTypeIDText = "select QuestionTypeID from PptQuestionEntity where PaperType=@papertype";
            //DataTable pptQuestionTypeIDDt;
            //pptQuestionTypeIDDt = sqlhelper.ExecuteQuery(pptQuestionTypeIDText, pptpapertype, CommandType.Text);
            //string pptQuestionTypeID;
            //pptQuestionTypeID = pptQuestionTypeIDDt.Rows[0]["QuestionTypeID"].ToString();

            //在pptSumFration表中插入数据
            DateTime dtime = System.DateTime.Now;
            SqlParameter[] parasinsertPPT = new SqlParameter[]{ new SqlParameter("@StudentID",studentinfo.studentID),new SqlParameter("@PaperType","NULL"),new SqlParameter("@QuestionTypeID","NULL"),new SqlParameter("@sumfartion",sumFration),new SqlParameter("@TimeStamp",dtime)
            };
            string insertppt = "insert into PptSumFration values(@StudentID,@PaperType,@QuestionTypeID,@sumfartion,@TimeStamp)";
            int numexcel = sqlhelper.ExecuteNonQuery(insertppt, parasinsertPPT, CommandType.Text);
            return pptsumFrationDt;
        }
        #endregion

        #region IE类汇总总分-李芬

        /// <summary>
        /// IE类汇总总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable IEsumfrationdal(StudentInfoEntity studentinfo)
        {
            String which = selectCollegeID(studentinfo).ToString();
            //查询学生的试卷类型/类型总分值
            SqlParameter[] IEparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string sumText = "select sum(convert(float,fraction)) from IEQuestionRecordEntity_" + which + " where StudentID=@studentID group by StudentID";

            DataTable IEsumFrationDt;
            IEsumFrationDt = sqlhelper.ExecuteQuery(sumText, IEparas, CommandType.Text);
            string sumFration;
            sumFration = IEsumFrationDt.Rows[0][0].ToString();



            //在IESumFration表中插入数据
            DateTime dtime = System.DateTime.Now;
            SqlParameter[] parasinsertIE = new SqlParameter[]{ new SqlParameter("@StudentID",studentinfo.studentID),new SqlParameter("@PaperType","NULL"),new SqlParameter("@QuestionTypeID","NULL"),new SqlParameter("@sumfartion",sumFration),new SqlParameter("@TimeStamp",dtime)
            };
            string insertIE = "insert into IESumFration values(@StudentID,@PaperType,@QuestionTypeID,@sumfartion,@TimeStamp)";
            int numie = sqlhelper.ExecuteNonQuery(insertIE, parasinsertIE, CommandType.Text);

            return IEsumFrationDt;

        }
        #endregion
        #region 选择题类汇总总分-李芬

        /// <summary>
        /// 选择题类汇总总分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable Selectsumfration(StudentInfoEntity studentinfo)
        {
            String which = selectCollegeID(studentinfo).ToString();
            //查询学生的试卷类型/类型总分值
            SqlParameter[] Selectparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string selectsumText = "select sum(convert(float,Fration)) from SelectQuestionRecordEntity_" + which + " where StudentID=@studentID group by StudentID";

            DataTable SelectsumFrationDt;
            SelectsumFrationDt = sqlhelper.ExecuteQuery(selectsumText, Selectparas, CommandType.Text);
            string SelectsumFration;
            SelectsumFration = SelectsumFrationDt.Rows[0][0].ToString();



            //在SelectQuestionRecordEntity表中插入数据
            DateTime dtime = System.DateTime.Now;
            SqlParameter[] parasinsertSelect = new SqlParameter[]{ new SqlParameter("@StudentID",studentinfo.studentID),new SqlParameter("@PaperType","NULL"),new SqlParameter("@QuestionTypeID","NULL"),new SqlParameter("@sumfartion",SelectsumFration),new SqlParameter("@TimeStamp",dtime)
            };
            string insertSelect = "insert into SelectSumFration values(@StudentID,@PaperType,@QuestionTypeID,@sumfartion,@TimeStamp)";
            int numSelect = sqlhelper.ExecuteNonQuery(insertSelect, parasinsertSelect, CommandType.Text);

            return SelectsumFrationDt;

        }
        #endregion

        #region 学生的总成绩汇总-李芬

        /// <summary>
        /// 学生的总成绩汇总
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable sumfrationdal(StudentInfoEntity studentinfo)
        {
            String which = selectCollegeID(studentinfo).ToString();

            string examID;
            string examPlaceID;


            #region 查询考场的id，考场地点名称
            //1.查询考场的id，考场地点名称
            DataTable examstudent = new DataTable();
            string examText = "select examID,examPlaceID from ExamPlaceEntity ";
            examstudent = sqlhelper.ExecuteQuery(examText, CommandType.Text);
            examID = examstudent.Rows[0]["examID"].ToString();
            examPlaceID = examstudent.Rows[0]["examPlaceID"].ToString();

            #endregion



            #region 查询学生的姓名-李芬
            //查询学生的姓名
            SqlParameter[] paras = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string studentnameText = "select studentName from StudentInfoEntity where StudentID=@studentID";

            DataTable studentNameDt;
            studentNameDt = sqlhelper.ExecuteQuery(studentnameText, paras, CommandType.Text);
            string studentName = studentNameDt.Rows[0]["studentName"].ToString();
            #endregion



            //查询word中的分数
            #region 查询word中的分数-李芬
            SqlParameter[] wordparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string wordsum = "select Fration from WordSumFration where StudentID=@studentID";

            DataTable wordsumDt;
            wordsumDt = sqlhelper.ExecuteQuery(wordsum, wordparas, CommandType.Text);
            string wordsumfration = wordsumDt.Rows[0]["Fration"].ToString();
            #endregion


            //查询win中的分数
            #region 查询win中的分数-李芬
            SqlParameter[] winparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string winsum = "select Fration from WinSumFration where StudentID=@studentID";

            DataTable winsumDt;
            winsumDt = sqlhelper.ExecuteQuery(winsum, winparas, CommandType.Text);
            string winsumfration = winsumDt.Rows[0]["Fration"].ToString();
            #endregion


            //查询选择题中的分数
            #region 查询选择题中的分数-李芬
            SqlParameter[] selectparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string selectsum = "select Fration from SelectSumFration where StudentID=@studentID";

            DataTable selectsumDt;
            selectsumDt = sqlhelper.ExecuteQuery(selectsum, selectparas, CommandType.Text);
            string selectsumfration = selectsumDt.Rows[0]["Fration"].ToString();
            #endregion

            //查询ppt中的分数
            #region 查询ppt中的分数-李芬
            SqlParameter[] pptparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string pptsum = "select Fration from PptSumFration where StudentID=@studentID";

            DataTable pptsumDt;
            pptsumDt = sqlhelper.ExecuteQuery(pptsum, pptparas, CommandType.Text);
            string pptsumfration = pptsumDt.Rows[0]["Fration"].ToString();
            #endregion


            //查询IE中的分数
            #region 查询IE中的分数-李芬
            SqlParameter[] IEparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string IEsum = "select Fration from IESumFration where StudentID=@studentID";

            DataTable IEsumDt;
            IEsumDt = sqlhelper.ExecuteQuery(IEsum, IEparas, CommandType.Text);
            string IEsumfration = IEsumDt.Rows[0]["Fration"].ToString();
            #endregion

            //#region 查询OutLook中的分数-李芬
            ////查询OutLook中的分数
            //SqlParameter[] OutLookparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            //string OutLooksum = "select Fration from OutLookSumFration where StudentID=@studentID";

            //DataTable OutLooksumDt;
            //OutLooksumDt = sqlhelper.ExecuteQuery(OutLooksum, OutLookparas, CommandType.Text);
            //string OutLooksumfration = OutLooksumDt.Rows[0]["Fration"].ToString();
            //#endregion


            #region 查询Excel中的分数-李芬
            //查询Excel中的分数
            SqlParameter[] Excelparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string Excelsum = "select Fration from ExcelSumFration where StudentID=@studentID";

            DataTable ExcelsumDt;
            ExcelsumDt = sqlhelper.ExecuteQuery(Excelsum, Excelparas, CommandType.Text);
            string Excelsumfration = ExcelsumDt.Rows[0]["Fration"].ToString();

            #endregion

            #region 在总分数的表中插入记录-李芬
            //在总分数的表中插入记录
            float sumFration = float.Parse(Excelsumfration) + float.Parse(IEsumfration) + float.Parse(pptsumfration) + float.Parse(selectsumfration) + float.Parse(winsumfration) + float.Parse(wordsumfration);

            //在总成绩表中插入记录，合并考场号、考场地点、个人总分数等信息
            SqlParameter[] parasinsert = new SqlParameter[]{new SqlParameter ("@studentID",studentinfo.studentID),
            new SqlParameter("@score",sumFration),new SqlParameter("@examID",examID),
            new SqlParameter("@examPlaceID",examPlaceID),new SqlParameter("@collegeID",which),new SqlParameter("@studentName",studentName)
                   };
            string insertText = "insert into ScoreEntity values(@examID,@examPlaceID,@studentID,@studentName,@score,@collegeID)";
            int num = sqlhelper.ExecuteNonQuery(insertText, parasinsert, CommandType.Text);
            #endregion


            return studentNameDt;

        }
        #endregion

        #region 查询学生的学院—李芬
        /// <summary>
        /// 查询学生的学院—李芬
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public string selectCollegeID(StudentInfoEntity studentinfo)
        {
            SqlParameter[] parascollegeID = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string collegeIDText = "select CollegeID from StudentInfoEntity where StudentID=@studentID";

            DataTable collegeDT;
            collegeDT = sqlhelper.ExecuteQuery(collegeIDText, parascollegeID, CommandType.Text);

            string SelectCollegeID;
            SelectCollegeID = collegeDT.Rows[0]["CollegeID"].ToString();

            return SelectCollegeID;

        }

        #endregion
        /// <summary>
        /// 查看是否已经判分
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable selectstudentID(StudentInfoEntity studentinfo)
        {
            SqlParameter[] studentIDparas = new SqlParameter[] { new SqlParameter("@studentID", studentinfo.studentID) };
            string studentIDText = "select StudentID from ScoreEntity where StudentID=@studentID";
            DataTable selectIDscore;
            selectIDscore = sqlhelper.ExecuteQuery(studentIDText, studentIDparas, CommandType.Text);
            return selectIDscore;

        }
    }
}
