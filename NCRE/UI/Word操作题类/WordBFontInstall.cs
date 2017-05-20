/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/17 11:54:08 
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
    public class WordBFontInstall

    {
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordB.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private WordQuestionEntity wordinfo = new WordQuestionEntity();

        #region 王荣晓 标题字体型号设计
        /// <summary>
        /// 字体设置：标题字体型号设计
        /// 
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontNameInstallB(WordQuestionEntity wordinfo)
        {
            
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体型号";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.Name.ToString();
            
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体型号和数据库中的一样
                if (doc.Paragraphs[2].Range.Font.Name == dt.Rows[i]["RightAnswer"].ToString())               
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

        #region 王荣晓 标题字体大小设计
        /// <summary>
        /// 字体设置：标题字体大小设计
        /// 
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontSizeInstallB(WordQuestionEntity wordinfo)
        {
    
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体大小";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.Size.ToString();

                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体型号和数据库中的一样
                if (doc.Paragraphs[2].Range.Font.Size == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 王荣晓 标题字体加粗设置
        /// <summary>
        /// 标题字体加粗设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontBoldInstallB(WordQuestionEntity wordinfo)
        {
            
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体加粗";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.Bold.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体是否粗体和数据库中的一样
                if (doc.Paragraphs[2].Range.Font.Bold.ToString () == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 标题字体颜色设置
        /// <summary>
        /// 标题字体颜色设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontColorInstallB(WordQuestionEntity wordinfo)
        {   

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体颜色";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.ColorIndex.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体颜色和数据库中的一样
                if (doc.Paragraphs[2].Range.Font.ColorIndex.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 标题字体间距
        /// <summary>
        /// 标题字体间距
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontSpaceInstallB(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体间距";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.Spacing.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体颜色和数据库中的一样
                if (doc.Paragraphs[2].Range.Font.Spacing.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 标题字体格式
        /// <summary>
        /// 标题字体格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontAlignInstallB(WordQuestionEntity wordinfo)
        {
            

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体格式";
            //调出根据问题标识和试卷类型查出的信息
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Alignment.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体对齐方式是否和数据库中的一样
                if (doc.Paragraphs[2].Alignment.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 正文字体型号设置

        /// <summary>
        /// 正文字体型号设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextSetB(WordQuestionEntity wordinfo)
        {
           
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文字体型号";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Font.Name.ToString();
                //doc.Paragraphs.Last.Range.Font.Size+
                //app.Selection.ParagraphFormat.Alignment.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体类型为数据库中对应的字体
                if (doc.Paragraphs[4].Range.Font.Name.ToString() == dt.Rows[i]["RightAnswer"].ToString())
                // doc.Paragraphs.Last.Range.Font.Size == 10.5 &
                //app.Selection.ParagraphFormat.Alignment.ToString() == "wdAlignParagraphCenter")
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

        #region 王荣晓 正文字体加粗
        /// <summary>
        /// 正文字体加粗
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextBoldSetB(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文字体加粗";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Font.BoldBi.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文字体大小和数据库中的一样
                if (doc.Paragraphs[4].Range.Font.BoldBi == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 王荣晓 正文字体大小
        /// <summary>
        /// 正文字体大小
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextSizeSetB(WordQuestionEntity wordinfo)
        {
           

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文字体大小";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Font.Size.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文字体大小和数据库中的一样
                if (doc.Paragraphs[4].Range.Font.Size == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 王荣晓 正文格式
        /// <summary>
        /// 正文格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextFormatSetB(WordQuestionEntity wordinfo)
        {
            

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文格式";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = app.Selection.ParagraphFormat.Alignment.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文格式和数据库中的一样
                if (app.Selection.ParagraphFormat.Alignment.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 正文行距
        /// <summary>
        /// 正文行距
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextLineSpacingB(WordQuestionEntity wordinfo)
        {         

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文行距";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].LineSpacing.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文格式和数据库中的一样
                if (doc.Paragraphs[4].LineSpacing.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
