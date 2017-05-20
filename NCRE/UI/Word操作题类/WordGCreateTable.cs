/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/12/9 16:08:02 
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
using System.Windows.Forms;


namespace NCRE学生考试端V1._0
{
    class WordGCreateTable
    {
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\bgG.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();

        #region 李少然 表格行高  2015年12月9日16:14:14
        /// <summary>
        /// 行高
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetOtherLineHeightG(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格行高";
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Height.ToString() + doc.Tables[1].Cell(2, 1).Height.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).Height.ToString() + doc.Tables[1].Cell(2, 1).Height.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然  表格列宽 2015年12月9日16:15:09
        /// <summary>
        /// 表格第一列列宽
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetFirstColWeightG(WordQuestionEntity wordinfo)
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Width.ToString() + doc.Tables[1].Cell(2, 1).Width.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).Width.ToString() + doc.Tables[1].Cell(2, 1).Width.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 表格文字 2015年12月9日16:17:09
        /// <summary>
        /// 表格文字
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableWordFormatG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
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
                    studentrecord.ExamAnswer = doc.Tables[1].Range.Text.Trim().Substring(0, 5);

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Range.Text.Trim().Substring(0, 5) == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 表格单行字体型号 2015年12月9日19:31:13
        /// <summary>
        /// 表格单行字体型号
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableSingleLineG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格单行字体型号";
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

        #region 李少然 表格单行字体大小 2015年12月9日19:31:13
        /// <summary>
        /// 表格单行字体大小
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableSingleSizeG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格单行字体大小";
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

        #region 李少然 表格单行字体颜色 2015年12月9日19:31:13
        /// <summary>
        /// 表格单行字体颜色
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableSingleSizeColorG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格单行字体颜色";
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Shading.ForegroundPatternColor.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).Shading.ForegroundPatternColor.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 表格字体型号 2015年12月9日19:31:13
        /// <summary>
        /// 表格字体型号
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableLineG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格字体型号";
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(2, 1).Range.Font.Name.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(2, 1).Range.Font.Name.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 表格字体大小 2015年12月9日19:31:13
        /// <summary>
        /// 表格字体大小
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableLineSizeG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格字体大小";
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(2, 1).Range.Font.Size.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(2, 1).Range.Font.Size.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 表格字体颜色 2015年12月9日19:31:13
        /// <summary>
        /// 表格字体颜色
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableLineSizeColorG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格字体颜色";
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(2, 1).Shading.ForegroundPatternColor.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(2, 1).Shading.ForegroundPatternColor.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region  李少然 表格外边框线 2015年12月9日16:19:32
        /// <summary>
        /// 表格外边框线
        /// </summary>
        /// <param name="studentinfo"></param>
        public void TableBoldBolderG(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格外边框线";
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

        #region  李少然 表格外边框线颜色 2015年12月9日16:21:54
        /// <summary>
        /// 表格外边框线颜色
        /// </summary>
        /// <param name="studentinfo"></param>
        public void TableBoldBolderOutColorG(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格外边框线颜色";
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
                    studentrecord.ExamAnswer = doc.Tables[1].Borders.OutsideColorIndex.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Borders.OutsideColorIndex.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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


        #region  李少然 表格内边框线 2015年12月9日16:22:31
        /// <summary>
        /// 表格内边框线
        /// </summary>
        /// <param name="studentinfo"></param>
        public void TableInsideBoldBolderG(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格内边框线";
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

        #region  李少然 表格内边框线颜色 2015年12月9日16:22:50
        /// <summary>
        /// 表格内边框线颜色
        /// </summary>
        /// <param name="studentinfo"></param>
        public void TableBoldBolderInColorG(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格内边框线颜色";
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
                    studentrecord.ExamAnswer = doc.Tables[1].Borders.InsideColorIndex.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Borders.InsideColorIndex.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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


        #region 李少然 表格文字格式 2015年12月9日16:17:09
        /// <summary>
        /// 表格文字格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableWordFontFormatG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格文字格式";
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).VerticalAlignment.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(1, 1).VerticalAlignment.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 表格格式 2015年12月9日16:17:09
        /// <summary>
        /// 表格格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableFormatG(WordQuestionEntity wordinfo)
        { //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格格式";
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

    }
}
