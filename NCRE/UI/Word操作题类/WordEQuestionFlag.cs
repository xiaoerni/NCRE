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
    class WordEQuestionFlag
    {
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();
        /// <summary>
        /// 判断是什么类型的题
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SwitchQuestionFlagE(WordQuestionEntity wordinfo)
        {
            WordEPageOperate wordpage = new WordEPageOperate();
            WordEFindKeyWord wordfindkeyword = new WordEFindKeyWord();
            WordEFontInstall wordfontinstall = new WordEFontInstall();
            WordECreateTable wordtable = new WordECreateTable();
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

                            wordfindkeyword.FindChinesebracketsE(wordinfo);
                            break;
                        case "查找空格":

                            wordfindkeyword.FindSearchSpacesE(wordinfo);
                            break;
                        case "查找空行":

                            wordfindkeyword.FindLineE(wordinfo);
                            break;
                        case "查找替换":

                            wordfindkeyword.FindReplacementE(wordinfo);
                            break;
                        case "查找图片":

                            wordfindkeyword.FindPictureE(wordinfo);
                            break;
                        case "图片宽高":

                            wordfindkeyword.FindPictureWidthE(wordinfo);
                            break;
                        case "标题字体型号":

                            wordfontinstall.FontNameInstallE(wordinfo);
                            break;
                        case "标题格式":
                            wordfontinstall.TitleRightIndentSetE(wordinfo);
                            break;
                        case "标题字体大小":
                            wordfontinstall.FontSizeInstallE(wordinfo);
                            break;
                        case "标题字体加粗":
                            wordfontinstall.FontBoldInstallE(wordinfo);
                            break;
                        case "标题文字底纹":
                            wordfontinstall.TiTleCaptionShadingE(wordinfo);
                            break;
                        case "标题段前段后":
                            wordfontinstall.FontParagraphInstallE(wordinfo);
                            break;
                        case "小标题字体型号":
                            wordfontinstall.LittleTitleSetE(wordinfo);
                            break;
                        case "小标题字体大小":
                            wordfontinstall.LittleSizeSetE(wordinfo);
                            break;
                        case "小标题字体加粗":
                            wordfontinstall.LittleBoldSetE(wordinfo);
                            break;
                        case "小标题字体格式":
                            wordfontinstall.LittleRightIndentSetE(wordinfo);
                            break;
                        case "小标题段前段后":
                            wordfontinstall.LittleAlignSetE(wordinfo);
                            break;
                        case "正文字体型号":
                            wordfontinstall.MainTextSetE(wordinfo);
                            break;
                        case "正文字体大小":
                            wordfontinstall.MainTextSizeSetE(wordinfo);
                            break;
                        case "正文格式":
                            wordfontinstall.MainTextFormatSetE(wordinfo);
                            break;
                        case "正文行距":
                            wordfontinstall.TextSpacingE(wordinfo);
                            break;
                        case "页边距上下":

                            wordpage.PageMarginUpOperateE(wordinfo);
                            break;
                        case "页边距左右":

                            wordpage.PageMarginLeftOperateE(wordinfo);
                            break;
                        case "页眉页脚边距":

                            wordpage.PageMarginTopOperateE(wordinfo);
                            break;
                        case "纸张大小":

                            wordpage.PageSizeE(wordinfo);
                            break;
                        case "页眉文字":

                            wordpage.PageHeaderTextE(wordinfo);
                            break;
                        case "页眉格式":

                            wordpage.PageHeaderFormatE(wordinfo);
                            break;
                        case "查找页码":

                            wordpage.SearchPageE(wordinfo);
                            break;                       
                        case "表格行高":

                            wordtable.SetOtherLineHeightE(wordinfo);
                            break;
                        case "表格列宽":
                            wordtable.SetFirstColWeightE(wordinfo);
                            break;
                        case "表格文字":
                            wordtable.TableWordFormatE(wordinfo);
                            break;
                        case "表格文字型号":
                            wordtable.TableWordNameE(wordinfo);
                            break;
                        case "表格文字大小":
                            wordtable.TableWordSizeE(wordinfo);
                            break;
                        case "表格外边框线":
                            wordtable.TableBoldBolderE(wordinfo);
                            break;

                        //++++++++++++++++buduide++++++++++++++++++++
                        case "表格外边框线颜色":
                            wordtable.TableBoldBolderOutColorE(wordinfo);
                            break;
                        case "表格内边框线":
                            wordtable.TableInsideBoldBolderE(wordinfo);
                            break;
                        case "表格内边框线颜色":
                            wordtable.TableBoldBolderInColorE(wordinfo);
                            break;
                       case "表格单行填充颜色":
                            wordtable.TableSingelInColorE(wordinfo);
                            break;
                            
                        case "表格文字格式":
                            wordtable.TableWordFontFormatE(wordinfo);
                            break;
                        case "表格格式":
                            wordtable.TableFormatE(wordinfo);
                            break;                       
                        default:
                            break;
                    }

                }
            }
        }
    }
}
