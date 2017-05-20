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
    class WordGQuestionFlag
    {
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();
        /// <summary>
        /// 判断是什么类型的题
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SwitchQuestionFlagG(WordQuestionEntity wordinfo)
        {
            WordGPageOperate wordpage = new WordGPageOperate();
            WordGFindKeyWord wordfindkeyword = new WordGFindKeyWord();
            WordGFontInstall wordfontinstall = new WordGFontInstall();
            WordGCreateTable wordtable = new WordGCreateTable();
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
                        case "查找文字":

                            wordfindkeyword.FindFontG(wordinfo);
                            break;
                        case "查找空行":

                            wordfindkeyword.FindSearchSpacesG(wordinfo);
                            break;                      
                        case "查找字体":

                            wordfindkeyword.FindFontShapeG(wordinfo);
                            break;
                        case "查找图片":

                            wordfindkeyword.FindPictureG(wordinfo);
                            break;
                        case "图片宽高":

                            wordfindkeyword.FindPictureWidthG(wordinfo);
                            break;
                        case "标题字体型号":

                            wordfontinstall.FontNameInstallG(wordinfo);
                            break;
                        case "标题字体格式":
                            wordfontinstall.TitleRightIndentSetG(wordinfo);
                            break;
                        case "标题字体大小":
                            wordfontinstall.FontSizeInstallG(wordinfo);
                            break;
                        case "标题字体加粗":
                            wordfontinstall.FontBoldInstallG(wordinfo);
                            break;                       
                        case "标题字体颜色":
                            wordfontinstall.FontColorG(wordinfo);
                            break;                        
                        case "正文字体型号":
                            wordfontinstall.MainTextSetG(wordinfo);
                            break;
                        case "正文字体大小":
                            wordfontinstall.MainTextSizeSetG(wordinfo);
                            break;
                        case "正文字体格式":
                            wordfontinstall.MainTextFormatSetG(wordinfo);
                            break;
                        case "正文行距":
                            wordfontinstall.TextSpacingG(wordinfo);
                            break;
                        case "页边距上下":

                            wordpage.PageMarginUpOperateG(wordinfo);
                            break;
                        case "页边距左右":

                            wordpage.PageMarginLeftOperateG(wordinfo);
                            break;
                        case "页眉页脚边距":

                            wordpage.PageMarginTopOperateG(wordinfo);
                            break;
                        case "纸张大小":

                            wordpage.PageSizeG(wordinfo);
                            break;
                        case "页眉文字":

                            wordpage.PageHeaderTextG(wordinfo);
                            break;
                        case "页眉字体格式":

                            wordpage.PageHeaderFormatG(wordinfo);
                            break;
                        case "页眉字体型号":

                            wordpage.PageFontG(wordinfo);
                            break;
                        case "页眉字体大小":

                            wordpage.PageFontSizeG(wordinfo);
                            break;
                        case "表格行高":

                            wordtable.SetOtherLineHeightG(wordinfo);
                            break;
                        case "表格列宽":
                            wordtable.SetFirstColWeightG(wordinfo);
                            break;
                        case "表格文字":
                            wordtable.TableWordFormatG(wordinfo);
                            break;
                        case "表格单行字体型号":
                            wordtable.TableSingleLineG(wordinfo);
                            break;
                        case "表格单行字体大小":
                            wordtable.TableSingleSizeG(wordinfo);
                            break;
                        case "表格单行字体颜色":
                            wordtable.TableSingleSizeColorG(wordinfo);
                            break;
                        case "表格字体型号":
                            wordtable.TableLineG(wordinfo);
                            break;
                        case "表格字体大小":
                            wordtable.TableLineSizeG(wordinfo);
                            break;
                        case "表格字体颜色":
                            wordtable.TableLineSizeColorG(wordinfo);
                            break;
                       
                        case "表格外边框线":
                            wordtable.TableBoldBolderG(wordinfo);
                            break;
                        case "表格外边框线颜色":
                            wordtable.TableBoldBolderOutColorG(wordinfo);
                            break;
                        case "表格内边框线":
                            wordtable.TableInsideBoldBolderG(wordinfo);
                            break;
                        case "表格内边框线颜色":
                            wordtable.TableBoldBolderInColorG(wordinfo);
                            break;

                        //++++++++++++++++buduide++++++++++++++++++++
                        
                        case "表格文字格式":
                            wordtable.TableWordFontFormatG(wordinfo);
                            break;
                        case "表格格式":
                            wordtable.TableFormatG(wordinfo);
                            break;                       
                        default:
                            break;
                    }

                }
            }
        }
    }
}
