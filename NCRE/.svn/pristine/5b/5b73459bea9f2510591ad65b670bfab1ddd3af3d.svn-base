/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/20 19:46:30 
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
    public class WordCQuestionFlag
    {
         
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private WordQuestionEntity wordinfo = new WordQuestionEntity();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();

        public void SwitchQuestionFlagC(WordQuestionEntity wordinfo)
        {
            WordCPageOperate wordpage = new WordCPageOperate();
            WordCFindKeyWord wordfindkeyword = new WordCFindKeyWord();
            WordCFontInstall wordfontinstall = new WordCFontInstall();
            WordCCreateTable wordtable = new WordCCreateTable();
            wordinfo.PaperType = "C";
             DataTable dt = wordquestionbll.LoadWordQuestion(wordinfo);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                   string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                String collegeId = MyInfo.MycollegeID();
                if (examconfigbll.IsTableExist("WordQuestionRecordEntity_" + collegeId) == false)
                {
                    break;
                }
                else
                {

                    switch (questionflag)
                    {
                        
                        case "删除空格":
                            wordfindkeyword.DeleteBlankWordC(wordinfo);
                            break;
                        
                        case "替换段落":
                            wordfindkeyword.ReplaceParagraphWordC(wordinfo);
                            break;
                        case "替换文字":
                            wordfindkeyword.FindKeyWordC(wordinfo);
                            break;

                        case "查找页码":
                            wordfontinstall.FindPageNumC(wordinfo);
                            break;
                        
                        case "查找图片":
                            wordfontinstall.FindPictureC(wordinfo);
                            break;
                        
                        case "图片高宽":
                            wordfontinstall.FindPictureHightC(wordinfo);
                            break;
                        case "标题字体型号":

                            wordfontinstall.FontNameInstallC(wordinfo);
                            break;
                        case "标题字体颜色":
                            wordfontinstall.FontColorInstallC(wordinfo);
                            break;
                        case "标题字体大小":
                            wordfontinstall.FontSizeInstallC(wordinfo);
                            break;
                        
                        case "标题字体加粗":
                            wordfontinstall.FontBoldInstallC(wordinfo);
                            break;
                        case "标题格式":
                            wordfontinstall.FontAlignInstallC(wordinfo);
                            break;
                        case "小标题格式":
                            wordfontinstall.LittleFontAlignInstallC(wordinfo);
                            break;

                        case "小标题字体型号":

                            wordfontinstall.LittleFontNameInstallC(wordinfo);
                            break;
                        case "小标题字体颜色":
                            wordfontinstall.LittleFontColorInstallC(wordinfo);
                            break;
                        case "小标题字体大小":
                            wordfontinstall.LittleFontSizeInstallC(wordinfo);
                            break;

                        case "小标题字体加粗":
                            wordfontinstall.LittleFontBoldInstallC(wordinfo);
                            break;


                        case "页边距":

                            wordpage.PageMarginUpOperateC(wordinfo);
                            break;
                        case "纸张大小":

                            wordpage.PageSizeC(wordinfo);
                            break;
                        case "小标题段前段后":
                            wordfontinstall.TittleAlignSetC(wordinfo);
                            break;

                        case "正文字体型号":
                            wordfontinstall.MainTextSetC(wordinfo);
                            break;
                        case "页眉文字":
                            wordfontinstall.HeaderTextSetC(wordinfo);
                            break;
                        case "页眉字体型号":
                            wordfontinstall.HeaderTextTypeSetC(wordinfo);
                            break;
                        case "页眉字体大小":
                            wordfontinstall.HeaderTextSizeSetC(wordinfo);
                            break;
                        case "页眉字体格式":
                            wordfontinstall.HeaderTextFormatSetC(wordinfo);
                            break;
                        case "页眉页脚边距":
                            wordfontinstall.HeaderMarginSetC(wordinfo);
                            break;
                        case "正文格式":
                            wordfontinstall.MainTextFormatSetC(wordinfo);
                            break;
                        case "正文行距":
                            wordfontinstall.MainTextLineSpacingC(wordinfo);
                            break;
                        case "表格格式":
                            wordtable.CreateTableC(wordinfo);
                            break;
                        case "表格列宽":
                            wordtable.SetFirstColWeightNC(wordinfo);
                            break;
                        case "表格行高":
                            wordtable.SetOtherLineHeightC(wordinfo);
                            break;
                        case "表格文字型号":
                            wordtable.FindCharFormatC(wordinfo);
                            break;
                        case "表格文字大小":
                            wordtable.FindCharFontSizeC(wordinfo);
                            break;
                        
                        case "表格文字加粗":
                            wordtable.FindCharFontBoldC(wordinfo);
                            break;
                        case "表格文字格式":
                            wordtable.TableWordFormatC(wordinfo);
                            break;

                        //++++++++++++++++buduide++++++++++++++++++++
                        case "表格外边框线":
                            wordtable.TableBoldBolderC(wordinfo);
                            break;
                        case "表格外边框线颜色":
                            wordtable.TableBoldBolderOutColorC(wordinfo);
                            break;
                        case "表格内边框线":
                            wordtable.TableInsideBoldBolderC(wordinfo);
                            break;
                        case "表格内边框线颜色":
                            wordtable.TableBoldBolderInColorC(wordinfo);
                            break;
                        
                        case "查找表格内容":
                            wordtable.FindFormTextC(wordinfo);
                            break;
                        default:
                            break;
                    }
                }
                }
            }
     }
    }

