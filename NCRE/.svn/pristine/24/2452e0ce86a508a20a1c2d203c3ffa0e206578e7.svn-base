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
    class WordHQuestionFlag
    {
        private WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();
        /// <summary>
        /// 判断是什么类型的题
        /// </summary>
        /// <param name="wordinfo"></param>
        public void SwitchQuestionFlagH(WordQuestionEntity wordinfo)
        {
            WordHPageOperate wordpage = new WordHPageOperate();
            WordHFindKeyWord wordfindkeyword = new WordHFindKeyWord();
            WordHFontInstall wordfontinstall = new WordHFontInstall();
            WordHCreateTable wordtable = new WordHCreateTable();
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
                        case "查找标点":

                            wordfindkeyword.FindPunctuationH(wordinfo);
                            break;
                        case "查找空行":

                            wordfindkeyword.FindSearchSpacesH(wordinfo);
                            break;                      
                        case "查找文字":

                            wordfindkeyword.FindFontShapeH(wordinfo);
                            break;
                        case "查找图片":

                            wordfindkeyword.FindPictureH(wordinfo);
                            break;
                        case "图片宽高":

                            wordfindkeyword.FindPictureWidthH(wordinfo);
                            break;
                        case "标题文字":

                            wordfontinstall.TiTleFontH(wordinfo);
                            break;
                        case "标题字体型号":

                            wordfontinstall.FontNameInstallH(wordinfo);
                            break;
                        case "标题字体格式":
                            wordfontinstall.TitleRightIndentSetH(wordinfo);
                            break;
                        case "标题字体大小":
                            wordfontinstall.FontSizeInstallH(wordinfo);
                            break;
                        case "标题字体加粗":
                            wordfontinstall.FontBoldInstallH(wordinfo);
                            break;                       
                        case "标题字体颜色":
                            wordfontinstall.FontColorH(wordinfo);
                            break;
                        case "标题段前段后":
                            wordfontinstall.TitleParagraphH(wordinfo);
                            break;
                        case "小标题字体型号":
                            wordfontinstall.LittleTitleSetH(wordinfo);
                            break;
                        case "小标题字体大小":
                            wordfontinstall.LittleSizeSetH(wordinfo);
                            break;
                        case "小标题字体加粗":
                            wordfontinstall.LittleBoldSetH(wordinfo);
                            break;
                        case "小标题字体格式":
                            wordfontinstall.LittleRightIndentSetH(wordinfo);
                            break;
                        case "小标题段前段后":
                            wordfontinstall.LittleAlignSetH(wordinfo);
                            break; 
                        case "正文字体型号":
                            wordfontinstall.MainTextSetH(wordinfo);
                            break;
                        case "正文字体大小":
                            wordfontinstall.MainTextSizeSetH(wordinfo);
                            break;
                        case "正文字体加粗":
                            wordfontinstall.MainTextBoldSetH(wordinfo);
                            break;
                        case "正文字体格式":
                            wordfontinstall.MainTextFormatSetH(wordinfo);
                            break;
                        case "正文行距":
                            wordfontinstall.TextSpacingH(wordinfo);
                            break;
                        case "页边距":

                            wordpage.PageMarginUpOperateH(wordinfo);
                            break;                       
                        case "页眉页脚边距":

                            wordpage.PageMarginTopOperateH(wordinfo);
                            break;
                        case "纸张大小":

                            wordpage.PageSizeH(wordinfo);
                            break;
                        case "页眉文字":

                            wordpage.PageHeaderTextH(wordinfo);
                            break;
                        case "页眉格式":

                            wordpage.PageHeaderFormatH(wordinfo);
                            break;
                        case "页眉字体型号":

                            wordpage.PageFontH(wordinfo);
                            break;
                        case "页眉字体大小":

                            wordpage.PageFontSizeH(wordinfo);
                            break;
                        case "页码格式":

                            wordpage.PageNumberHeaderFormatH(wordinfo);
                            break;
                        case "表格行高":

                            wordtable.SetOtherLineHeightH(wordinfo);
                            break;
                        case "表格列宽":
                            wordtable.SetFirstColWeightH(wordinfo);
                            break;
                        case "表格文字":
                            wordtable.TableWordFormatH(wordinfo);
                            break;                      
                        case "表格单行底纹":
                            wordtable.TableSingleShadingH(wordinfo);
                            break;
                        case "表格单行文字颜色":
                            wordtable.TableLineSizeColorH(wordinfo);
                            break;
                        case "表格字体型号":
                            wordtable.TableLineH(wordinfo);
                            break;
                        case "表格字体大小":
                            wordtable.TableLineSizeH(wordinfo);
                            break;
                        case "表格外边框线":
                            wordtable.TableBoldBolderH(wordinfo);
                            break;
                        case "表格外边框线颜色":
                            wordtable.TableBoldBolderOutColorH(wordinfo);
                            break;
                        case "表格内边框线":
                            wordtable.TableInsideBoldBolderH(wordinfo);
                            break;
                        case "表格内边框线颜色":
                            wordtable.TableBoldBolderInColorH(wordinfo);
                            break;

                        //++++++++++++++++buduide++++++++++++++++++++
                        
                        case "表格文字格式":
                            wordtable.TableWordFontFormatH(wordinfo);
                            break;
                        case "表格格式":
                            wordtable.TableFormatH(wordinfo);
                            break;                       
                        default:
                            break;
                    }

                }
            }
        }
    }
}
