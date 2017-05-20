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
    public  class WordAQuestionFlag
    {
        
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private WordQuestionEntity wordinfo = new WordQuestionEntity();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();

        /// <summary>
        /// 判断是什么类型的题
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SwitchQuestionFlag(WordQuestionEntity wordinfo)
        {
            WordAPageOperate wordpage = new WordAPageOperate();
            WordAFindKeyWord wordfindkeyword = new WordAFindKeyWord();
            WordAFontInstall wordfontinstall = new WordAFontInstall();
            WordACreateTable wordtable = new WordACreateTable();
          //根据试卷类型查找题型
            DataTable dt = wordquestionbll.LoadWordQuestion(wordinfo);
            //判断是否动态建库成功
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //添加list，放答题记录



                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                String collegeId = MyInfo.MycollegeID();
                //判断是否学生记录表存在
                if (examconfigbll.IsTableExist("WordQuestionRecordEntity_" + collegeId) == false)
                {
                    break;
                }
                else { 
               //成功了就开始判分
                    switch (questionflag)
                    {
                        case "插入空行":

                            //得到一个答题记录的实体

                            wordfindkeyword.FindKeyWord(wordinfo);
                            break;

                        case "查找替换":
                   
                            wordfindkeyword.FindKeyWordA(wordinfo);
                            break;
                        case "标题字体型号":
                   
                            wordfontinstall.FontNameInstall (wordinfo);
                            break;
                        case "标题字体颜色":
                            wordfontinstall.FontColorInstall(wordinfo);
                            break;
                        case "标题字体大小":
                            wordfontinstall.FontSizeInstall(wordinfo);
                            break;
                        case "标题字体加粗":
                            wordfontinstall.FontBoldInstall(wordinfo);
                            break;
                        case "标题字体对齐方式":
                            wordfontinstall.FontAlignInstall(wordinfo);
                            break;
                        
                        case "标题字符间距":
                            wordfontinstall.FontSeparationInstall(wordinfo);
                            break;
                        case "页边距上下":
                            wordpage.PageMarginUpOperate(wordinfo);
                            break;
                            
                        case "页边距左右":

                            wordpage.PageMarginLeftOperate(wordinfo);
                            break;
                        case "装订线位置":
                            wordpage.GutterPosition(wordinfo);
                            break;
                        case "纸张方向":
                            wordpage.PageDirection(wordinfo);
                            break;
                        case "纸张大小":
                            wordpage.PageSize(wordinfo);
                            break;
                        
                        case "查找页码":
                            wordpage.FindPageNum(wordinfo);
                            break;
                        
                        case "查找图片":
                            wordpage.FindPicture(wordinfo);
                            break;
                        case "图片高度":
                            wordpage.PictureWidth(wordinfo);
                            break;
                        case "正文字体型号":
                 
                            wordfontinstall.MainTextSet(wordinfo);
                            break;
                        case "正文字体大小":
              
                            wordfontinstall.MainTextSizeSet(wordinfo);
                            break;
                        case "正文格式":
                 
                            wordfontinstall.MainTextFormatSet(wordinfo);
                            break;
                        
                        case "正文行距":

                            wordfontinstall.SpacingFormatSet(wordinfo);
                            break;
                        case "表格行高":
                            wordtable.SetLineHeight(wordinfo);
                            break;
                        case "表格列宽":
                            wordtable.SetColWeight(wordinfo);
                            break;

                            //++++++++++++++++buduide++++++++++++++++++++
                        case "合并单元格文字":
                            wordtable.SetFirstColFont(wordinfo);
                            break;
                        case "表格外边框线":
                            wordtable.FormBorderFont(wordinfo);
                            break;
                        case "表格内边框线":
                            wordtable.FormInBorderFont(wordinfo);
                            break;
                        case "表格边框线颜色":
                            wordtable.FormInBorderColorFont(wordinfo);
                            break;
                        
                        case "表格底纹":
                            wordtable.SetColShading(wordinfo);
                            break;
                        case "表格格式":
                            wordtable.TableWordFormatA(wordinfo);
                            break;
                    
                        case "表格文字格式":
                            wordtable.TableWordFormat(wordinfo);
                            break;

                        //++++++++++++++++buduide++++++++++++++++++++
                        case "单行表格字体型号":
                            wordtable.SetOneRowFontName(wordinfo);
                            break;
                        
                        case "单行表格字体大小":
                            wordtable.SetOneRowFontSize(wordinfo);
                            break;
                        case "表格文字型号":
                            wordtable.SetFontName(wordinfo);
                            break;
                        case "表格文字大小":
                            wordtable.SetFontSize(wordinfo);
                            break;
                        case "表格字体加粗":
                            wordtable.SetFontBold(wordinfo);
                            break;

                        default:
                            break;
                    }
                    //更新list，插入记录
            }
            }
        }
    }
}



