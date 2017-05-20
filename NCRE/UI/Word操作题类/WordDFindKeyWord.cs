/********************************************************************************** 
     * 开发人:王荣晓
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/21 08:45:17 
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
    public class WordDFindKeyWord
    {
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordD.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);

        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();

        #region  王荣晓 查找替换
        /// <summary>
        /// 查找替换
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindKeyWordD(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找替换";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Content.Text.Trim().Substring(139, 3);
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Content.Text.Trim().Substring(139, 3) == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
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


        #region  王荣晓  查找互换
        /// <summary>
        /// 查找替换
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindParagraphD(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找互换";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Content.Text.Trim().Substring(173, 7);
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Content.Text.Trim().Substring(173, 7) == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
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

        #region  王荣晓  删除空行
        /// <summary>
        /// 查找替换
        /// </summary>
        /// <param name="studentinfo"></param>
        public void DeleteNullStringD(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "删除空行";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();
            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Content.Text.Trim().Substring(4, 1);
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();
                //将得分传递到studentrecord实体
                if (doc.Content.Text.Trim().Substring(4, 1) == dt.Rows[i]["RightAnswer"].ToString())//如果已经替换成功
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

    }
}
