/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/20 17:17:40 
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
using System.Threading;

namespace NCRE学生考试端V1._0
{
    public class WordCCreateTable

    {
        
        private static Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        private static object unknow = Type.Missing;       
        private static object file = @"D:\计算机一级考生文件\Wordkt\bgC.docx";
        private Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(ref file, ref unknow, ref unknow, ref unknow);
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
      

        #region 王荣晓 插入表格，新建表格
        /// <summary>
        /// 新建表格
        /// </summary>
        /// <param name="wordinfo"></param>
        public void CreateTableC(WordQuestionEntity wordinfo)
        {
            WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
           //判断表格是否存在            
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
            //判断表格是否存在          
                //循环遍历正确答案进行判分
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //传递考生ID进studentrecord实体
                    studentrecord.StudentID = FrmLogin.studentID;

                    //将考生答案传递给studentrecord实体
                    studentrecord.ExamAnswer = doc.Tables[1].Rows.Count.ToString() +
                        doc.Tables[1].Columns.Count.ToString();
                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Rows.Count.ToString() +
                           doc.Tables[1].Columns.Count.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
        /// 表格第一列列宽
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetFirstColWeightNC(WordQuestionEntity wordinfo)
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(3, 2).Width.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(3, 2).Width.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 表格行高
        /// <summary>
        /// 行高
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SetOtherLineHeightC(WordQuestionEntity wordinfo)
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(3, 1).Height.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(3, 1).Height.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region 王荣晓 表格文字格式
        //表格文字格式
        public void TableWordFormatC(WordQuestionEntity wordinfo)
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
                    studentrecord.ExamAnswer = doc.Tables[1].Cell(3, 2).VerticalAlignment.ToString();

                    //将试题的ID选择出来
                    studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                    //将每道题的分值取出
                    fration = dt.Rows[i]["Fration"].ToString();
                    //将得分传递到studentrecord实体

                    if (doc.Tables[1].Cell(3, 2).VerticalAlignment.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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

        #region  王荣晓 表格文字型号
        /// <summary>
        /// 表格文字型号
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindCharFormatC(WordQuestionEntity wordinfo)
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

        #region  王荣晓 表格文字大小
        /// <summary>
        /// 表格文字大小
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindCharFontSizeC(WordQuestionEntity wordinfo)
        {
            //将正确答案、分值取出来，传给dt
            wordinfo.QuestionFlag = "表格文字大小";
            DataTable dt = wordquestionbll.LoadWordByFlag(wordinfo);
            WordQuestionRecordEntity   studentrecord = new WordQuestionRecordEntity  ();

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

        #region  王荣晓 表格文字加粗
        /// <summary>
        /// 表格文字加粗
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindCharFontBoldC(WordQuestionEntity wordinfo)
        {

             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格文字加粗";
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
                     studentrecord.ExamAnswer = doc.Tables[1].Cell(1, 1).Range.Font.BoldBi.ToString();

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Cell(1, 1).Range.Font.BoldBi.ToString() == dt.Rows[i]["RightAnswer"].ToString())
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
        
        #region  王荣晓 表格外边框线
        /// <summary>
        /// 表格外边框线
        /// </summary>
        /// <param name="studentinfo"></param>
        public void TableBoldBolderC(WordQuestionEntity wordinfo)
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

        #region 王荣晓 表格外边框线颜色
        /// <summary>
        /// 表格外边框线颜色
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableBoldBolderOutColorC(WordQuestionEntity wordinfo)
        {

             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
              //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格外边框线颜色";
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

        #region 王荣晓 表格内边框线
        /// <summary>
        /// 表格内边框线
        /// </summary>
        /// <param name="wordinfo"></param>
        public void TableInsideBoldBolderC(WordQuestionEntity wordinfo)
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

        #region 王荣晓 表格内边框线颜色
        public void TableBoldBolderInColorC(WordQuestionEntity wordinfo)
        {

             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "表格内边框线颜色";
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

        #region  王荣晓 查找表格内容
        /// <summary>
        /// 查找表格内容
        /// </summary>
        /// <param name="studentinfo"></param>
        public void FindFormTextC(WordQuestionEntity wordinfo)
        {
             WordQuestionRecordEntity studentrecord = new WordQuestionRecordEntity();
            
                 //将正确答案、分值取出来，传给dt
                 wordinfo.QuestionFlag = "查找表格内容";
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
                     studentrecord.ExamAnswer = doc.Tables[1].Cell(5, 2).Range.Text.Substring(0, 2);

                     //将试题的ID选择出来
                     studentrecord.QuestionID = Convert.ToDouble(dt.Rows[i]["QuestionID"]);
                     //将每道题的分值取出
                     fration = dt.Rows[i]["Fration"].ToString();
                     //将得分传递到studentrecord实体

                     if (doc.Tables[1].Cell(5, 2).Range.Text.Substring (0,2) == dt.Rows[i]["RightAnswer"].ToString())
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
