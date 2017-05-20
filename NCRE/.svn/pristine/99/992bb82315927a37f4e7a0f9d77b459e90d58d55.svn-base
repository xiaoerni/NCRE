/********************************************************************************** 
     * 开发人:王荣晓
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/21 08:23:40 
     *开发版本：V1.0
 **********************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using Model;
using System.Data;
using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;


namespace NCRE学生考试端V1._0
{
    public class WordDCreateTable
    {
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\bgD.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();

        #region 王荣晓  表格列宽
        /// <summary>
        /// 表格列宽
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetFirstColWeightND(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格列宽";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {
               
                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Width.ToString() + doc.Tables[1].Cell(1, 3).Width.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).Width.ToString() + doc.Tables[1].Cell(1, 3).Width.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }

        }
        #endregion

        #region 王荣晓 表格行高
        /// <summary>
        /// 行高
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetOtherLineHeightD(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格行高";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(3, 1).Height.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(3, 1).Height.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }

        }
        #endregion

        #region  王荣晓 表格文字
        /// <summary>
        /// 表格文字
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindFormTextD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格文字";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Range.Text.Substring(0, 2);

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).Range.Text.Substring (0,2) == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }
        }

        #endregion

        #region  王荣晓 表格文字格式
        /// <summary>
        /// 表格文字格式  
        /// </summary>
        /// <param name="studentinfo"></param>
        public void WordFormD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格文字格式  ";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(3, 2).VerticalAlignment.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(3, 2).VerticalAlignment.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }
        }

        #endregion

        #region  王荣晓 表格格式
        /// <summary>
        /// 表格格式  
        /// </summary>
        /// <param name="studentinfo"></param>
        public void WordFormTypeD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格格式  ";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Rows.Alignment.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Rows.Alignment.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }
        }

        #endregion

        #region  王荣晓 表格文字型号
        /// <summary>
        /// 表格文字型号
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindCharFormatD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格文字型号";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Range.Font.Name.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).Range.Font.Name.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }
        }

        #endregion

        #region  王荣晓 表格文字大小
        /// <summary>
        /// 表格文字大小
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindCharFontSizeD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格文字大小";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Range.Font.Size.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).Range.Font.Size.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }
        }

        #endregion

        #region  王荣晓 表格文字加粗
        /// <summary>
        /// 表格文字加粗
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindCharFontBoldD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格文字加粗";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Range.Font.BoldBi.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).Range.Font.BoldBi.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }
        }

        #endregion

        #region  王荣晓 表格外边框线
        /// <summary>
        /// 表格外边框线
        /// </summary>
        /// <param name="studentinfo"></param>
        public void TableBoldBolderD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格外边框线";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Borders.OutsideLineWidth.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Borders.OutsideLineWidth.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }
        }

        #endregion

        #region 王荣晓 表格内边框线
        /// <summary>
        /// 表格内边框线
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableInsideBoldBolderD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格内边框线";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            if (doc.Tables.Count == 0)
            {
                //如果表格不存在，则判分为0
                studentrecord.StudentID = FrmLogin.studentID;
                studentrecord.ExamAnswer = "0";
                studentrecord.Fration = "0";
                //将答题记录送到数据库
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                }
                wordquestionbll.ReturnScore(studentrecord);

            }
            else
            {

                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Borders.InsideLineWidth.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Borders.InsideLineWidth.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                    {
                        //加分
                        studentrecord.Fration = fration;
                    }
                    else
                    {
                        //不加分
                        studentrecord.Fration = "0";
                    }

                    wordquestionbll.ReturnScore(studentrecord);
                }
                return;
            }
        }
        #endregion


    }
}
