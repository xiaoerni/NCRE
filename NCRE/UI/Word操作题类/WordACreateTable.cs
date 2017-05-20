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

   public class WordACreateTable
    {
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;
        private static object file = @"D:\计算机一级考生文件\Wordkt\bgA.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private WordQuestionEntity wordinfo = new WordQuestionEntity();


        #region 王荣晓  表格行高
        /// <summary>
        /// 表格行高
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetLineHeight(WordQuestionEntity wordinfo)
        {
          
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();           
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格行高";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
                string fration;
                if (doc.Tables .Count == 0)
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
                     studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 2).Height.ToString() + doc.Tables[1].Cell(2, 2).Height.ToString();

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Cell(1, 2).Height.ToString() + doc.Tables[1].Cell(2, 2).Height.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓  表格列宽
        /// <summary>
        ///  表格列宽
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetColWeight(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格列宽";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
                string fration;
                if (doc.Tables .Count == 0)
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
                     studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Width.ToString() + doc.Tables[1].Cell(1, 2).Width.ToString();

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Cell(1, 1).Width.ToString() + doc.Tables[1].Cell(1, 2).Width.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓  合并单元格文字
        /// <summary>
        /// 合并单元格文字   
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetFirstColFont(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();             
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "合并单元格文字";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
                string fration;
                if (doc.Tables .Count == 0)
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
                     studentrecord.ExamAnswer = doc.Tables[1].Range.Text.Trim().Substring(0, 6);

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Range.Text.Trim().Substring(0, 6) == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓  表格外边框线
        /// <summary>
        /// 表格外边框线   
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FormBorderFont(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格外边框线";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
               string fration;
               if (doc.Tables .Count == 0)
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

        #region 王荣晓  表格内边框线
        /// <summary>
        /// 表格内边框线   
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FormInBorderFont(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格内边框线";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
                string fration;
                if (doc.Tables .Count == 0)
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

        #region 王荣晓  表格边框线颜色
        /// <summary>
        /// 表格边框线颜色   
        /// </summary>
        /// <param name="wordinfo"></param>
        public void FormInBorderColorFont(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格边框线颜色";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);          
                 string fration;
                 if (doc.Tables .Count == 0)
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

        #region 王荣晓  表格底纹
        /// <summary>
        ///  表格底纹
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetColShading(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格底纹";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
               string fration;
               if (doc.Tables .Count == 0)
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
                     studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Shading.BackgroundPatternColorIndex.ToString();

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Cell(1, 1).Shading.BackgroundPatternColorIndex.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓  表格文字格式
        /// <summary>
        ///  表格文字格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableWordFormat(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格文字格式";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
                string fration;
                if (doc.Tables .Count == 0)
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

        #region 王荣晓  表格格式
        /// <summary>
        ///  表格格式
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableWordFormatA(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格格式";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
                string fration;
                if (doc.Tables .Count == 0)
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

        #region 王荣晓  单行表格字体型号
        /// <summary>
        /// 单行表格字体型号   
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetOneRowFontName(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "单行表格字体型号";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
                string fration;
                if (doc.Tables .Count == 0)
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

                     if (doc.Tables[1].Cell(1, 1).Range.Font.Name == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓  单行表格字体大小
        /// <summary>
        /// 单行表格字体大小   
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetOneRowFontSize(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
           
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "单行表格字体大小";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
               string fration;
               if (doc.Tables .Count == 0)
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

        #region 王荣晓 表格字体型号
        /// <summary>
        /// 表格字体型号     
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetFontName(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格文字型号";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);             
                 string fration;
                 if (doc.Tables .Count == 0)
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
                     studentrecord.ExamAnswer = doc.Tables[1].Cell(2, 2).Range.Font.Name.ToString();

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Cell(2, 2).Range.Font.Name.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region  王荣晓  表格字体大小
        /// <summary>
        /// 表格字体大小     
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetFontSize(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格文字大小";
                 DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
               string fration;
               if (doc.Tables .Count == 0)
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
                     studentrecord.ExamAnswer = doc.Tables[1].Cell(2, 2).Range.Font.Size.ToString();

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Cell(2, 2).Range.Font.Size.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region  王荣晓  表格字体加粗
        /// <summary>
        /// 表格字体加粗     
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetFontBold(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
             //将正确答案、分值取出来，传给dt
             wordinfo.QuestionFlag = "表格字体加粗";
              DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
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
                     studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Range.Font.BoldBi.ToString ();

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Cell(1, 1).Range.Font.BoldBi.ToString () == dt.Rows[i]["RightAnswer"].ToString())
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

