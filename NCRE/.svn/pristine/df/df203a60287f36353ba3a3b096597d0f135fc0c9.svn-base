/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 周洲、李少然、陈晓婵、王虹芸、李芬、王荣晓
     * 类说明：  
     * 开发时间：2015/12/9 14:47:58 
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
    class WordHFindKeyWord
    {

        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordH.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();

        #region  李少然 查找标点  2015年12月9日19:12:07
        /// <summary>
        /// 查找标点
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void FindPunctuationH(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找标点";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Text.Trim().Substring(0, 4);
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Paragraphs[1].Range.Text.Trim().Substring(0, 4) == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
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
        #endregion        

        #region  李少然 查找空行  2015年12月9日19:12:07
        /// <summary>
        /// 查找空行
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void FindSearchSpacesH(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找空行";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Text.Trim().Substring(0, 4);
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Paragraphs[1].Range.Text.Trim().Substring(0,4) == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
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
        #endregion        

        #region  李少然 查找文字  2015年12月9日19:13:40
        /// <summary>
        /// 查找文字
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void FindFontShapeH(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找文字";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Text.Trim().Substring(0, 4);
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Paragraphs[1].Range.Text.Trim().Substring(0, 4) == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
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
        #endregion        

        

        #region  李少然 查找图片  2015年12月9日16:06:21
        /// <summary>:10
        /// 查找图片
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void FindPictureH(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找图片";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Shapes.Count.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Shapes.Count.ToString() == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
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
        #endregion

        #region  李少然 图片宽高  2015年12月9日16:06:21
        /// <summary>:10
        /// 图片宽高
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void FindPictureWidthH(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "图片宽高";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            string fration;
            if (doc.Shapes.Count == 0)
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
                    studentrecord.ExamAnswer = doc.Shapes[1].Height.ToString() + doc.Shapes[1].Width.ToString();
                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体
                    if (doc.Shapes[1].Height.ToString() + doc.Shapes[1].Width.ToString() == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
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
