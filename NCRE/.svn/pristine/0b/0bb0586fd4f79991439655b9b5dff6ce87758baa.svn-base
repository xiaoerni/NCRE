/********************************************************************************** 
     * 开发人:李少然
     * 开发组： 
     * 类说明：  
     * 开发时间：2015/11/18 9:14:13 
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
    public class WordBQuestionFlag
    {
        
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        private WordQuestionEntity wordinfo = new WordQuestionEntity();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();

        public void SwitchQuestionFlagB(WordQuestionEntity wordinfo)
        {
            WordBPageOperate wordpage = new WordBPageOperate();
            WordBFindKeyWord wordfindkeyword = new WordBFindKeyWord();
            WordBFontInstall wordfontinstall = new WordBFontInstall();
            WordBCreateTable wordtable = new WordBCreateTable();

                wordinfo.PaperType = "B";
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
                        case "插入空行":
                            wordfindkeyword.FindKeyWordB(wordinfo);
                            break;
                        case "查找替换":
                            wordfindkeyword.FindReplaceWordB(wordinfo);
                            break;
                        case "装订线位置":
                            wordpage.GutterPositionB(wordinfo);
                            break;
                        case "纸张方向":
                            wordpage.PageDirectionB(wordinfo);
                            break;
                        case "纸张大小":
                            wordpage.PageSizeB(wordinfo);
                            break;

                        case "查找页码":
                            wordpage.FindPageNumB(wordinfo);
                            break;

                        case "查找图片":
                            wordpage.FindPictureB(wordinfo);
                            break;
                        
                        case "图片高度":
                            wordpage.FindPictureHightB(wordinfo);
                            break;
                        case "标题字体型号":

                            wordfontinstall.FontNameInstallB(wordinfo);
                            break;
                        case "标题字体颜色":
                            wordfontinstall.FontColorInstallB(wordinfo);
                            break;
                        case "标题字体大小":
                            wordfontinstall.FontSizeInstallB(wordinfo);
                            break;
                        case "标题字体加粗":
                            wordfontinstall.FontBoldInstallB(wordinfo);
                            break;
                        
                        case "标题字体间距":
                            wordfontinstall.FontSpaceInstallB(wordinfo);
                            break;
                        case "标题字体格式":
                            wordfontinstall.FontAlignInstallB(wordinfo);
                            break;
                        case "页边距上下":

                            wordpage.PageMarginUpOperateB(wordinfo);
                            break;
                        case "页边距左右":

                            wordpage.PageMarginLeftOperateB(wordinfo);
                            break;
                        
                        case "正文字体型号":

                            wordfontinstall.MainTextSetB(wordinfo);
                            break;
                        case "正文字体大小":

                            wordfontinstall.MainTextSizeSetB(wordinfo);
                            break;
                        case "正文字体加粗":

                            wordfontinstall.MainTextBoldSetB(wordinfo);
                            break;
                        case "正文格式":

                            wordfontinstall.MainTextFormatSetB(wordinfo);
                            break;
                        case "正文行距":

                            wordfontinstall.MainTextLineSpacingB(wordinfo);
                            break;
                        case "表格列宽":
                            wordtable.SetColWeightB(wordinfo);
                            break;
                        case "表格行高":
                            wordtable.SetLineHeightB(wordinfo);
                            break;

                        case "表格格式":
                            wordtable.TableFormatB(wordinfo);
                            break;
                        case "表格文字格式":
                            wordtable.TableWordFormatB(wordinfo);
                            break;

                        //++++++++++++++++buduide++++++++++++++++++++
                        case "表格外边框线":
                            wordtable.FormBorderFontB(wordinfo);
                            break;
                        case "表格边框线颜色":
                            wordtable.FormInBorderColorFontB(wordinfo);
                            break;
                        case "表格内边框线":
                            wordtable.FormInBorderFontB(wordinfo);
                            break;
                        default:
                            break;
                    }
                }
                }
            }
     }
    
}
