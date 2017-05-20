/********************************************************************************** 
     * 开发人:王荣晓
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/21 08:57:30 
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
    public class WordDQuestionFlag
    {
        
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private WordQuestionEntity wordinfo = new WordQuestionEntity();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();


        public void SwitchQuestionFlagD(WordQuestionEntity wordinfo)
        {
            WordDPageOperate wordpage = new WordDPageOperate();
            WordDFindKeyWord wordfindkeyword = new WordDFindKeyWord();
            WordDFontInstall wordfontinstall = new WordDFontInstall();
            WordDCreateTable wordtable = new WordDCreateTable();
            wordinfo.PaperType = "D";
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
                        case "查找替换":
                            wordfindkeyword.FindKeyWordD(wordinfo);
                            break;
                        case "查找互换":
                            wordfindkeyword.FindParagraphD(wordinfo);
                            break;
                        case "删除空行":
                            wordfindkeyword.DeleteNullStringD(wordinfo);
                            break;

                        case "标题字体型号":
                            wordfontinstall.FontNameInstallD(wordinfo);
                            break;
                        case "标题字体颜色":
                            wordfontinstall.FontColorInstallD(wordinfo);
                            break;
                        case "标题字体大小":
                            wordfontinstall.FontSizeInstallD(wordinfo);
                            break;
                        
                        case "标题字体加粗":
                            wordfontinstall.FontBoldInstallD(wordinfo);
                            break;
                        case "标题字体格式":
                            wordfontinstall.FontAlignInstallD(wordinfo);
                            break;

                        case "页边距上下":
                            wordpage.PageMarginUpOperateD(wordinfo);
                            break;
                        case "页边距左右":
                            wordpage.PageMarginLeftOperateD(wordinfo);
                            break;
                        case "纸张大小":
                            wordpage.PageSizeD(wordinfo);
                            break;
                        case "页眉页脚边距":
                            wordpage.HeaderFootOperateD(wordinfo);
                            break;

                        case "正文字体型号":

                            wordfontinstall.MainTextSetD(wordinfo);
                            break;
                        case "正文字体大小":

                            wordfontinstall.MainTextSizeSetD(wordinfo);
                            break;
                        case "正文格式":

                            wordfontinstall.MainTextFormatSetD(wordinfo);
                            break;
                        case "小标题字体加粗":
                            wordfontinstall.LittleTextFontBoldInstallD(wordinfo);
                            break;
                        
                        case "小标题字体型号":
                            wordfontinstall.LittleTextSetD(wordinfo);
                            break;
                        
                        case "小标题格式":
                            wordfontinstall.LittleTextFormatSetD(wordinfo);
                            break;
                        case "小标题字体大小":
                            wordfontinstall.LittleSizeSetD(wordinfo);
                            break;
                        case "小标题段前段后":

                            wordfontinstall.LittleAlignSetD(wordinfo);
                            break;
                        case "小标题字体颜色":

                            wordfontinstall.LittleColorSetD(wordinfo);
                            break;
                        case "正文行距":

                            wordfontinstall.MainTextLineSpacingD(wordinfo);
                            break;
                        
                        case "正文字体加粗":

                            wordfontinstall.MainTextFontBoldInstallD(wordinfo);
                            break;

                        case "页眉文字":
                            wordfontinstall.HeaderTextSetD(wordinfo);
                            break;
                        case "页眉字体型号":
                            wordfontinstall.HeaderTextTypeSetD(wordinfo);
                            break;
                        case "页眉字体大小":
                            wordfontinstall.HeaderTextSizeSetD(wordinfo);
                            break;
                        case "页眉字体格式":
                            wordfontinstall.HeaderTextFormatSetD(wordinfo);
                            break;
                        case "页脚文字":
                            wordfontinstall.FindPageNumD(wordinfo);
                            break;
                        
                        case "页脚字体型号":
                            wordfontinstall.FindPageNumNameD(wordinfo);
                            break;
                        
                        case "查找图片":
                            wordfontinstall.FindPictureD(wordinfo);
                            break;

                        case "图片宽高":
                            wordfontinstall.FindPictureHightD(wordinfo);
                            break;
                        case "页脚字体大小":
                            wordfontinstall.FindPageNumSizeD(wordinfo);
                            break;
                        case "表格列宽":
                            wordtable.SetFirstColWeightND(wordinfo);
                            break;
                        case "表格行高":
                            wordtable.SetOtherLineHeightD(wordinfo);
                            break;
                        
                        case "表格文字":
                            wordtable.FindFormTextD(wordinfo);
                            break;
                        case "表格文字型号":
                            wordtable.FindCharFormatD(wordinfo);
                            break;
                        case "表格文字大小":
                            wordtable.FindCharFontSizeD(wordinfo);
                            break;
                        
                        case "表格文字加粗":
                            wordtable.FindCharFontBoldD(wordinfo);
                            break;
                        case "表格格式":
                            wordtable.WordFormTypeD(wordinfo);
                            break;
                        case "表格文字格式":
                            wordtable.WordFormD(wordinfo);
                            break;

                        case "表格外边框线":
                            wordtable.TableBoldBolderD(wordinfo);
                            break;
                        case "表格内边框线":
                            wordtable.TableInsideBoldBolderD(wordinfo);
                            break;
                        default:
                            break;
                    }
                }
            }
        }
    }
}
