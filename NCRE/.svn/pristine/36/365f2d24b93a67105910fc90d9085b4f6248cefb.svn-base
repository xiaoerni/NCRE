/********************************************************************************** 
     * 开发人:王荣晓
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/20 08:57:26 
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
    public class WordDPageOperate
    {
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordD.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);

        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();

        #region 王荣晓 纸张大小
        /// <summary>
        /// 纸张大小
        /// </summary>
        /// <param name="wordinfo"></param>
        public void PageSizeD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "纸张大小";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的上下边距边距传递给studentrecord实体               
                studentrecord.ExamAnswer = doc.PageSetup.PaperSize.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);

                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的上下边距和数据库中的一样
                if (doc.PageSetup.PaperSize.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 上下页边距
        /// <summary>
        /// 上下页边距操作
        /// </summary>
        /// <param name="wordinfo"></param>
        public void PageMarginUpOperateD(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页边距上下";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的上下边距边距传递给studentrecord实体               
                studentrecord.ExamAnswer = doc.PageSetup.TopMargin.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的上下边距和数据库中的一样
                if (doc.PageSetup.TopMargin.ToString () == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 左右页边距
        /// <summary>
        /// 左右页边距操作
        /// </summary>
        /// <param name="wordinfo"></param>
        public void PageMarginLeftOperateD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页边距左右";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的左右边距边距传递给studentrecord实体               
                studentrecord.ExamAnswer = doc.PageSetup.LeftMargin.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的左右边距和数据库中的一样
                if (doc.PageSetup.LeftMargin.ToString () == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 页眉页脚边距
        /// <summary>
        /// 页眉页脚边距
        /// </summary>
        /// <param name="wordinfo"></param>
        public void HeaderFootOperateD(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉页脚边距";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的左右边距边距传递给studentrecord实体               
                studentrecord.ExamAnswer = app.ActiveDocument.PageSetup.HeaderDistance.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(app.ActiveDocument.PageSetup.HeaderDistance.ToString());
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的左右边距和数据库中的一样
                if (app.ActiveDocument.PageSetup.HeaderDistance.ToString () == dt.Rows[i]["RightAnswer"].ToString())
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
