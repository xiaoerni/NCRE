/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/12/9 16:25:56 
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
    class WordFQuestionFlag
    {
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();
        /// <summary>
        /// 判断是什么类型的题
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SwitchQuestionFlagF(WordQuestionEntity wordinfo)
        {
            WordFPageOperate wordpage = new WordFPageOperate();
            WordFFindKeyWord wordfindkeyword = new WordFFindKeyWord();
            WordFFontInstall wordfontinstall = new WordFFontInstall();
            WordFCreateTable wordtable = new WordFCreateTable();
            //根据试卷类型查找题型
            DataTable dt = wordquestionbll.LoadWordQuestion(wordinfo);
            //判断是否动态建库成功
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                String collegeId = MyInfo.MycollegeID();
                //判断是否学生记录表存在
                if (examconfigbll.IsTableExist("WordQuestionRecordEntity_" + collegeId) == false)
                {
                    break;
                }
                else
                {
                    //成功了就开始判分
                    switch (questionflag)
                    {
                        case "查找括号":

                            wordfindkeyword.FindChinesebracketsF(wordinfo);
                            break;
                        case "查找空格":

                            wordfindkeyword.FindSearchSpacesF(wordinfo);
                            break;                      
                        case "查找替换":

                            wordfindkeyword.FindReplacementF(wordinfo);
                            break;
                        case "插入图片":

                            wordfindkeyword.FindPictureF(wordinfo);
                            break;
                        case "图片宽度":

                            wordfindkeyword.FindPictureWidthF(wordinfo);
                            break;
                        case "标题字体型号":

                            wordfontinstall.FontNameInstallF(wordinfo);
                            break;
                        case "标题格式":
                            wordfontinstall.TitleRightIndentSetF(wordinfo);
                            break;
                        case "标题字体大小":
                            wordfontinstall.FontSizeInstallF(wordinfo);
                            break;
                        case "标题字体加粗":
                            wordfontinstall.FontBoldInstallF(wordinfo);
                            break;                       
                        case "标题段前段后":
                            wordfontinstall.FontParagraphInstallF(wordinfo);
                            break;
                        case "小标题字体型号":
                            wordfontinstall.LittleTitleSetF(wordinfo);
                            break;
                        case "小标题字体大小":
                            wordfontinstall.LittleSizeSetF(wordinfo);
                            break;
                        case "小标题字体加粗":
                            wordfontinstall.LittleBoldSetF(wordinfo);
                            break;
                        case "小标题格式":
                            wordfontinstall.LittleRightIndentSetF(wordinfo);
                            break;
                        case "小标题段前段后":
                            wordfontinstall.LittleAlignSetF(wordinfo);
                            break;
                        case "正文字体型号":
                            wordfontinstall.MainTextSetF(wordinfo);
                            break;
                        case "正文字体大小":
                            wordfontinstall.MainTextSizeSetF(wordinfo);
                            break;
                        case "正文字体格式":
                            wordfontinstall.MainTextFormatSetF(wordinfo);
                            break;
                        case "正文行距":
                            wordfontinstall.TextSpacingF(wordinfo);
                            break;
                        case "页边距上下":

                            wordpage.PageMarginUpOperateF(wordinfo);
                            break;
                        case "页边距左右":

                            wordpage.PageMarginLeftOperateF(wordinfo);
                            break;                      
                        case "纸张大小":

                            wordpage.PageSizeF(wordinfo);
                            break;
                        case "页眉文字":

                            wordpage.PageHeaderTextF(wordinfo);
                            break;
                        case "页眉字体格式":

                            wordpage.PageHeaderFormatF(wordinfo);
                            break;
                        case "页码格式":

                            wordpage.SearchPageF(wordinfo);
                            break;                       
                        case "表格行高":

                            wordtable.SetOtherLineHeightF(wordinfo);
                            break;
                        case "表格列宽":
                            wordtable.SetFirstColWeightF(wordinfo);
                            break;
                        case "表格文字":
                            wordtable.TableWordFormatF(wordinfo);
                            break;
                       
                        case "表格外边框":
                            wordtable.TableBoldBolderF(wordinfo);
                            break;

                        //++++++++++++++++buduide++++++++++++++++++++
                        
                        case "表格文字格式":
                            wordtable.TableWordFontFormatF(wordinfo);
                            break;
                        case "表格格式":
                            wordtable.TableFormatF(wordinfo);
                            break;                       
                        default:
                            break;
                    }

                }
            }
        }
    }
}
