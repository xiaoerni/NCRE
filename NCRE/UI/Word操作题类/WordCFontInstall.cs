/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/20 16:43:20 
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
using Microsoft.Office.Interop.Word;

namespace NCRE学生考试端V1._0
{
   public  class WordCFontInstall
    {
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordC.docx";
        private static Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private static object oMissing = System.Reflection.Missing.Value;
        Range allRange = doc.Range(oMissing, oMissing); 

        #region 王荣晓 标题字体大小设计
        /// <summary>
        /// 字体设置：标题字体大小设计
        /// 
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontSizeInstallC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体大小";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Font.Size.ToString();

                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体型号和数据库中的一样
                if (doc.Paragraphs[1].Range.Font.Size == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 王荣晓 标题字体型号设计
        /// <summary>
        /// 字体设置：标题字体型号设计
        /// 
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontNameInstallC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体型号";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Font.Name.ToString();

                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体型号和数据库中的一样
                if (doc.Paragraphs[1].Range.Font.Name == dt.Rows[i]["RightAnswer"].ToString())
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
        public void FontBoldInstallC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体加粗";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Font.Bold.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体是否粗体和数据库中的一样
                if (doc.Paragraphs[1].Range.Font.Bold == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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
        public void FontColorInstallC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题字体颜色";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Font.ColorIndex.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体颜色和数据库中的一样
                if (doc.Paragraphs[1].Range.Font.ColorIndex.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
        public void FontAlignInstallC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题格式";
            //调出根据问题标识和试卷类型查出的信息
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = app.Selection.ParagraphFormat.Alignment.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体对齐方式是否和数据库中的一样
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



        #region 王荣晓 小标题字体大小设计
        /// <summary>
        /// 字体设置：小标题字体大小设计
        /// 
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleFontSizeInstallC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题字体大小";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Font.Size.ToString();

                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体型号和数据库中的一样
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

        #region 王荣晓 小标题字体型号设计
        /// <summary>
        /// 字体设置：小标题字体型号设计
        /// 
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleFontNameInstallC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题字体型号";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Font.Name.ToString();

                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体型号和数据库中的一样
                if (doc.Paragraphs[4].Range.Font.Name == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 小标题字体加粗设置
        /// <summary>
        /// 小标题字体加粗设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleFontBoldInstallC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题字体加粗";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Font.Bold.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体是否粗体和数据库中的一样
                if (doc.Paragraphs[4].Range.Font.Bold == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 王荣晓 小标题字体颜色设置
        /// <summary>
        /// 小标题字体颜色设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleFontColorInstallC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题字体颜色";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Font.ColorIndex.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体颜色和数据库中的一样
                if (doc.Paragraphs[4].Range.Font.ColorIndex.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 小标题字体对齐方式设置
        /// <summary>
        /// 小标题字体对齐方式设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleFontAlignInstallC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题格式";
            //调出根据问题标识和试卷类型查出的信息
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Alignment.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体对齐方式是否和数据库中的一样
                if (doc.Paragraphs[4].Alignment.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 小标题段前段后
        /// <summary>
        /// 小标题段前段后
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TittleAlignSetC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题段前段后";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(doc.Paragraphs[3].SpaceBefore.ToString());
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.ParagraphFormat.SpaceBefore.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果小标题段前段后和数据库中的相等
                if (doc.Paragraphs[4].Range.ParagraphFormat.SpaceBefore.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
        public void MainTextFormatSetC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文格式";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].FirstLineIndent.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文格式和数据库中的一样
                if (doc.Paragraphs[2].FirstLineIndent.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 正文字体型号

        /// <summary>
        /// 正文字体型号
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextSetC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文字体型号";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.NameFarEast.ToString() + doc.Paragraphs[2].Range.Font.NameAscii.ToString();
                //doc.Paragraphs.Last.Range.Font.Size+
                //app.Selection.ParagraphFormat.Alignment.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体类型为数据库中对应的字体
                if (doc.Paragraphs[2].Range.Font.NameFarEast.ToString() + doc.Paragraphs[2].Range.Font.NameAscii.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 正文行距
        /// <summary>
        /// 正文行距
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextLineSpacingC(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "正文行距";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[3].LineSpacing.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文格式和数据库中的一样
                if (doc.Paragraphs[3].LineSpacing.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
        /// 页眉文字
        /// </summary>
        /// <param name="wordinfo"></param>
        public void HeaderMarginSetC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉页脚边距";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;

            if (allRange.Sections[1].Headers.Count  == 0)
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
                    //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = app.ActiveDocument.PageSetup.HeaderDistance.ToString();
                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();

                    //将得分传递到studentrecord实体
                    //如果正文字体大小和数据库中的一样
                    if (app.ActiveDocument.PageSetup.HeaderDistance.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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


        #region 王荣晓 页眉文字
        /// <summary>
        /// 页眉文字
        /// </summary>
        /// <param name="wordinfo"></param>
        public void HeaderTextSetC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉文字";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

            string fration;

            if (allRange.Sections[1].Headers.Count  == 0)
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
                    //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.ToString();
                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();

                    //将得分传递到studentrecord实体
                    //如果正文字体大小和数据库中的一样
                    if (allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Trim ().ToString () == dt.Rows[i]["RightAnswer"].ToString())
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
        
        #region 王荣晓 页眉字体型号
        /// <summary>
        /// 页眉字体型号
        /// </summary>
        /// <param name="wordinfo"></param>
        public void HeaderTextTypeSetC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉字体型号";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文字体大小和数据库中的一样
                if (allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name.Trim ().ToString ()== dt.Rows[i]["RightAnswer"].ToString())
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
        
        #region 王荣晓 页眉字体大小
        /// <summary>
        /// 页眉字体大小
        /// </summary>
        /// <param name="wordinfo"></param>
        public void HeaderTextSizeSetC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉字体大小";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文字体大小和数据库中的一样
                if (allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
       
        #region 王荣晓 页眉字体格式
        /// <summary>
        /// 页眉字体格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void HeaderTextFormatSetC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉字体格式";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(app.Selection.ParagraphFormat.Alignment.ToString());

                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = app.ActiveWindow.View.SeekView.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文字体大小和数据库中的一样
                if (app.ActiveWindow.View.SeekView.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 查找页码
        /// <summary>
        /// 查找页码
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FindPageNumC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找页码";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的上下边距边距传递给studentrecord实体               
                studentrecord.ExamAnswer = doc.Footnotes.StartingNumber.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);

                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的上下边距和数据库中的一样
                if (doc.Footnotes.StartingNumber.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 查找图片
        /// <summary>
        /// 查找图片
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FindPictureC(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "查找图片";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的上下边距边距传递给studentrecord实体               
                studentrecord.ExamAnswer = doc.Shapes.Count.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);

                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的上下边距和数据库中的一样
                if (doc.Shapes.Count.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 图片高度
        /// <summary>
        /// 图片高度
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FindPictureHightC(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "图片高宽";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
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
                    //将考生答案的上下边距边距传递给studentrecord实体               
                    studentrecord.ExamAnswer = doc.Shapes[1].Height.ToString() + doc.Shapes[1].Width.ToString();
                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);

                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();

                    //将得分传递到studentrecord实体
                    //如果页的上下边距和数据库中的一样
                    if (doc.Shapes[1].Height.ToString() + doc.Shapes[1].Width.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
