/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/12/9 15:12:56 
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
    class WordHPageOperate
    {

        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
         private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\WordH.docx";
         private static Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private static object oMissing = System.Reflection.Missing.Value;
        Range allRange = doc.Range(oMissing, oMissing); 

        #region 李少然 页边距  2015年12月9日19:16:47
      /// <summary>
      /// 页边距
      /// </summary>
      /// <param name="wordinfo"></param>
        public void PageMarginUpOperateH(WordQuestionEntity wordinfo)
        {

            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页边距";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的上下边距边距传递给studentrecord实体               
                studentrecord.ExamAnswer = doc.PageSetup.TopMargin.ToString() + doc.PageSetup.BottomMargin.ToString() + doc.PageSetup.LeftMargin.ToString() + doc.PageSetup.RightMargin.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的上下边距和数据库中的一样
                if (doc.PageSetup.TopMargin.ToString() + doc.PageSetup.BottomMargin.ToString() + doc.PageSetup.LeftMargin.ToString() + doc.PageSetup.RightMargin.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 纸张大小
        /// <summary>
        /// 上下页边距操作
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void PageSizeH(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "纸张大小";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

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

        #region 李少然 页眉页脚边距 2015年12月9日15:17:17
        /// <summary>
        /// 页眉页脚边距
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void PageMarginTopOperateH(WordQuestionEntity wordinfo)
        {


            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉页脚边距";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的页眉页脚边距边距传递给studentrecord实体 
                studentrecord.ExamAnswer = app.ActiveDocument.PageSetup.HeaderDistance.ToString();
                // studentrecord.ExamAnswer = app.Selection.HeaderFooter.Range.ParagraphFormat.Alignment.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的页眉页脚边距和数据库中的一样
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
        #endregion

        #region 李少然 页眉文字  2015年12月9日15:52:28
        /// <summary>
        /// 页眉文字
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void PageHeaderTextH(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉文字";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            if (allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text == "\r")
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
                    //将考生答案的页眉页脚边距边距传递给studentrecord实体 
                    studentrecord.ExamAnswer = allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Substring(0, 4);
                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();

                    //将得分传递到studentrecord实体
                    //如果页的页眉页脚边距和数据库中的一样
                    if (allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Substring(0, 4) == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 页眉字体型号  2015年12月9日15:52:28
        /// <summary>
        /// 页眉字体型号
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void PageFontH(WordQuestionEntity wordinfo)
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
                //将考生答案的页眉页脚边距边距传递给studentrecord实体 
                studentrecord.ExamAnswer = allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的页眉页脚边距和数据库中的一样
                if (allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 李少然 页眉字体大小  2015年12月9日19:19:56
        /// <summary>
        /// 页眉字体大小
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void PageFontSizeH(WordQuestionEntity wordinfo)
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
                //将考生答案的页眉页脚边距边距传递给studentrecord实体 
                studentrecord.ExamAnswer = allRange.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的页眉页脚边距和数据库中的一样
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

        #region 李少然 页眉格式  2015年12月9日15:58:35
        /// <summary>
        /// 页眉格式
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void PageHeaderFormatH(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页眉格式";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的页眉页脚边距边距传递给studentrecord实体 
                studentrecord.ExamAnswer = app.ActiveWindow.View.SeekView.ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的页眉页脚边距和数据库中的一样
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

        #region 李少然 页码格式  2015年12月9日15:58:35
        /// <summary>
        /// 页码格式
        /// </summary>
        /// <param name="wordinfo">题库实体</param>
        public void PageNumberHeaderFormatH(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "页码格式";
            System.Data.DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();

            string fration;
            //循环遍历正确答案进行判分
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //传递考生ID进studentrecord实体
                studentrecord.StudentID = FrmLogin.studentID;
                //将考生答案的页眉页脚边距边距传递给studentrecord实体 
                studentrecord.ExamAnswer = allRange.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Trim().ToString();
                //将试题的ID选择出来
                studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                //将每道题的分值取出
                fration = dt.Rows[i]["Fration"].ToString();

                //将得分传递到studentrecord实体
                //如果页的页眉页脚边距和数据库中的一样
                if (allRange.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text.Trim().ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
