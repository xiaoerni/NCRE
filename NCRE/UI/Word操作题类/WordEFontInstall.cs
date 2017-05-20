/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/12/9 15:18:51 
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
    class WordEFontInstall
    {
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordE.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();

        #region 李少然 标题字体型号设计 2015年12月9日15:20:07
        /// <summary>
        /// 字体设置：标题字体型号设计
        /// 
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void FontNameInstallE(WordQuestionEntity wordinfo)
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
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Font.Name.ToString();            
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体型号和数据库中的一样
                if (doc.Paragraphs[1].Range.Font.Name.ToString() == dt.Rows[i]["RightAnswer"].ToString())             
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

        #region 李少然 标题字体加粗设置  2015年12月9日15:21:27
        /// <summary>
        /// 标题字体加粗设置
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void FontBoldInstallE(WordQuestionEntity wordinfo)
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
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Font.Bold.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体是否粗体和数据库中的一样
                if (doc.Paragraphs[1].Range.Font.Bold.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 标题字体大小设置  2015年12月9日15:22:32
        /// <summary>
        /// 标题字体大小设置
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void FontSizeInstallE(WordQuestionEntity wordinfo)
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
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Font.Size.ToString();
                //将试题的ID选择出来                
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体大小和数据库中的一样
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

        #region 李少然 标题格式 2015年12月9日15:23:48
        /// <summary>
        /// 标题段格式
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void TitleRightIndentSetE(WordQuestionEntity wordinfo)
        {
          //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题格式";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(doc.Paragraphs[3].SpaceBefore.ToString());
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = app.Selection.ParagraphFormat.Alignment.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果小标题格式和数据库中的相等
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

        #region 李少然 标题段前段后设置  2015年12月9日15:24:45
        /// <summary>
        /// 标题段前段后设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FontParagraphInstallE(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题段前段后";
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
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.ParagraphFormat.SpaceAfter.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果标题段前段后设置和数据库中的一样
                if (doc.Paragraphs[1].Range.ParagraphFormat.SpaceAfter == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 李少然 标题文字底纹  2015年12月9日15:24:45
        /// <summary>
        /// 标题文字底纹
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TiTleCaptionShadingE(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "标题文字底纹";
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
                studentrecord.ExamAnswer = doc.Paragraphs[1].Range.Font.UnderlineColor.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果标题段前段后设置和数据库中的一样
                if (doc.Paragraphs[1].Range.Font.UnderlineColor.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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



        #region 李少然 小标题字体型号设置 2015年12月9日15:31:19
        /// <summary>
        /// 小标题字体型号设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleTitleSetE(WordQuestionEntity wordinfo)
        {
           //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题字体型号";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
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
                //如果字体型号和数据库中的相等
                if (doc.Paragraphs[4].Range.Font.Name.ToString() == dt.Rows[i]["RightAnswer"].ToString())
               
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

        #region 李少然 小标题字体大小设置 2015年12月9日15:33:03
        /// <summary>
        /// 小标题字体大小设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleSizeSetE(WordQuestionEntity wordinfo)
        {
           //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题字体大小";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(doc.Paragraphs[3].SpaceBefore.ToString());
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Font.Size.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果小标题字体大小和数据库中的相等
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

        #region 李少然 小标题字体加粗设置
        /// <summary>
        /// 小标题字体加粗设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleBoldSetE(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题字体加粗";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(doc.Paragraphs[3].SpaceBefore.ToString());
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.Bold.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果小标题字体加粗和数据库中的相等
                if (doc.Paragraphs[4].Range.Bold == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 李少然 小标题格式  2015年12月9日15:33:58
        /// <summary>
        /// 小标题段格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleRightIndentSetE(WordQuestionEntity wordinfo)
        {
         //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题字体格式";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(doc.Paragraphs[3].SpaceBefore.ToString());
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].RightIndent.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果小标题格式和数据库中的相等
                if (doc.Paragraphs[4].RightIndent == float.Parse(dt.Rows[i]["RightAnswer"].ToString()))
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

        #region 李少然 小标题段前段后  2015年12月9日15:35:00
        /// <summary>
        /// 小标题段前段后
        /// </summary>
        /// <param name="wordinfo"></param>
        public void LittleAlignSetE(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "小标题段前段后";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //MessageBox.Show(doc.Paragraphs[3].SpaceBefore.ToString());
                //将考生答案传递给studentrecord实体
                studentrecord.ExamAnswer = doc.Paragraphs[4].Range.ParagraphFormat.SpaceAfter.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果小标题段前段后和数据库中的相等
                if (doc.Paragraphs[4].Range.ParagraphFormat.SpaceAfter.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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




        #region 李少然 正文字体型号设置  2015年12月9日15:36:32

        /// <summary>
        /// 正文字体型号设置
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextSetE(WordQuestionEntity wordinfo)
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
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.NameAscii.ToString() + doc.Paragraphs[2].Range.Font.NameFarEast.ToString();               
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果字体类型为数据库中对应的字体
                if (doc.Paragraphs[2].Range.Font.NameAscii.ToString() + doc.Paragraphs[2].Range.Font.NameFarEast.ToString() == dt.Rows[i]["RightAnswer"].ToString())              
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

        #region 李少然 正文字体大小  2015年12月9日15:38:28
        /// <summary>
        /// 正文字体大小
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextSizeSetE(WordQuestionEntity wordinfo)
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
                studentrecord.ExamAnswer = doc.Paragraphs[2].Range.Font.Size.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文字体大小和数据库中的一样
                if (doc.Paragraphs[2].Range.Font.Size.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 正文格式  2015年12月9日15:39:20
        /// <summary>
        /// 正文格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void MainTextFormatSetE(WordQuestionEntity wordinfo)
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

        #region 李少然 正文行距  2015年12月9日15:39:20
        /// <summary>
        /// 正文行距
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TextSpacingE(WordQuestionEntity wordinfo)
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
                studentrecord.ExamAnswer = doc.Paragraphs[2].LineSpacing.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果正文格式和数据库中的一样
                if (doc.Paragraphs[2].LineSpacing.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
