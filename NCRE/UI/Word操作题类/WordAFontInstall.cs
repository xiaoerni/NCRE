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

    public   class WordAFontInstall
    {

        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordA.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private WordQuestionEntity wordinfo = new WordQuestionEntity();


        #region 王荣晓 标题字体型号设计
        /// <summary>
        /// 字体设置：标题字体型号设计
        /// 
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontNameInstall(WordQuestionEntity wordinfo)
        {
            

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体型号";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.Name.ToString();

                //  +
                //  +
                // doc.Paragraphs[1].Range.Font.ColorIndex.ToString() +
                // doc.Paragraphs[1].Range.Font.Bold +
                //doc.Paragraphs[1].Range.ParagraphFormat.SpaceAfter.ToString() +
                // app.Selection.ParagraphFormat.Alignment ;
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体型号和数据库中的一样
                if (doc.Paragraphs[2].Range.Font.Name == dt.Rows[i]["RightAnswer"].ToString())
                // & 
                // doc.Paragraphs [1].Range.Font.Size==22 &
                // doc.Paragraphs[1].Range.Font.ColorIndex == Word.WdColorIndex.wdRed &
                // doc.Paragraphs[1].Range.Font.Bold == -1 &
                //doc.Paragraphs[1].Range.ParagraphFormat.SpaceAfter.ToString() == "62.4" &
                // app.Selection.ParagraphFormat.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter)
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
        public void FontColorInstall(WordQuestionEntity wordinfo)
        {
           
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体颜色";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

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

        #region 王荣晓 标题字体加粗设置
        /// <summary>
        /// 标题字体加粗设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontBoldInstall(WordQuestionEntity wordinfo)
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
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体是否粗体和数据库中的一样
                if (doc.Paragraphs[2].Range.Font.Bold == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 王荣晓 标题字体大小设置
        /// <summary>
        /// 标题字体大小设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontSizeInstall(WordQuestionEntity wordinfo)
        {
           
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体大小";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

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
                //如果字体大小和数据库中的一样
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

        #region 王荣晓 标题字符间距
        /// <summary>
        /// 标题字符间距
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontSeparationInstall(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字符间距";
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
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.Spacing.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体对齐方式是否和数据库中的一样
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

        #region 王荣晓 标题字体对齐方式设置
        /// <summary>
        /// 标题字体对齐方式设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontAlignInstall(WordQuestionEntity wordinfo)
        {
         

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体对齐方式";
            //调出根据问题标识和试卷类型查出的信息
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

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
        public void MainTextSet(WordQuestionEntity wordinfo)
        {
          

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文字体型号";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

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

        #region 王荣晓 正文字体大小
        /// <summary>
        /// 正文字体大小
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextSizeSet(WordQuestionEntity wordinfo)
        {
           

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文字体大小";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

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
        public void MainTextFormatSet(WordQuestionEntity wordinfo)
        {
          

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文格式";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

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
                studentrecord.QuestionID = Convert.ToDouble (dt.Rows[i]["QuestionID"]) ;
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
        public void SpacingFormatSet(WordQuestionEntity wordinfo)
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
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
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
   