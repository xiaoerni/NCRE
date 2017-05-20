using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using BLL;
using Model;
using System.Reflection;
using MSExcel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.IO;
using System.Data.SqlClient;
using NCRE学生考试端V1;
using System.Runtime.InteropServices;//互动服务  
using IWshRuntimeLibrary;
using System.Threading;

namespace NCRE学生考试端V1._0
{
    public class ExcelJudgeHelper
    {
        public static MSExcel.Application m_excel = null;            
        public static object unknow = null;
        public static string file = @"D:\计算机一级考生文件\Excelkt\Excel" + MyInfo.MyPaperType().Trim() + ".xlsx";
        public static MSExcel.Workbook m_workbook = null;
        public ExcelEntityBLL excelquestionbll =null;
        public ExcelQuestionEntity excelinfo = null;

        #region 根据题型关键字，判断判分方式——王虹芸
        /// <summary>
        /// 判断是什么类型的题
        /// </summary>
        /// <param name="excelinfo"></param>
        public void SelectJudge(ExcelQuestionEntity excelinfo)
        {
            if (System.IO.File.Exists(file)==false )
            {
                ExcelEntityBLL excelentitybll = new ExcelEntityBLL();
                ExcelQuestionRecordEntity excelrecord = new ExcelQuestionRecordEntity();
                excelrecord.PaperType = MyInfo.MyPaperType();
                excelrecord.StudentID = MyInfo.MystudentID();
                excelentitybll.UpdateExcelTypeID(excelrecord);
            }
            else
            {
                //调用Excel文件，并打开进行判分
                m_excel = new MSExcel.Application();
                unknow = Type.Missing;
                m_workbook = m_excel.Workbooks.Open(file,
                        unknow, unknow, unknow, unknow, unknow,
                        unknow, unknow, unknow, unknow, unknow,
                        unknow, unknow, unknow, unknow);
                excelquestionbll = new ExcelEntityBLL();
                excelinfo = new ExcelQuestionEntity();
                
                //开始判分
                ExcelQuestionEntity exceltype = new ExcelQuestionEntity();
                exceltype.PaperType = MyInfo.MyPaperType();
                DataTable dt = excelquestionbll.QueryExcelQuestionType(exceltype);
                int i;
                for (i = 0; i < dt.Rows.Count; i++)
                {
                    string qustionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                    //根据试题类型关键字判断
                    excelinfo.QuestionFlag = qustionflag;
                    excelinfo.PaperType = MyInfo.MyPaperType();
                    switch (qustionflag)
                    {
                        //修改名称来着的
                        case "工作表重命名":
                            ExcelSheetName excelsheetname = new ExcelSheetName();
                            excelsheetname.SheetName(excelinfo);
                            break;
                        case "图表类型":
                            ExcelChartType excelcharttype = new ExcelChartType();
                            excelcharttype.ChartType(excelinfo);
                            break;
                        case "图表标题":
                            ExcelChartTitle excelcharttitle = new ExcelChartTitle();
                            excelcharttitle.ChartTitle(excelinfo);
                            break;
                        case "图表颜色":
                            ExcelChartColor excelchartcolor = new ExcelChartColor();
                            excelchartcolor.ChartColor(excelinfo);
                            break;
                        case "图表位置":
                            ExcelChartPosition excelchartposition = new ExcelChartPosition();
                            excelchartposition.ChartPosition(excelinfo);
                            break;
                        case "图表标题样式":
                            ExcelChartTitleFont excelcharttitlefont = new ExcelChartTitleFont();
                            excelcharttitlefont.ChartTitleFont(excelinfo);
                            break;
                        case "图表标题大小":
                            ExcelChartTitleFontSize excelcharttitlefontsize = new ExcelChartTitleFontSize();
                            excelcharttitlefontsize.ChartTitleFontSize(excelinfo);
                            break;
                        case "图表标题颜色":
                            ExcelChartTitleFontColor excelcharttitlefontcolor = new ExcelChartTitleFontColor();
                            excelcharttitlefontcolor.ChartTitleFontColor(excelinfo);
                            break;
                        case "文件夹查找":
                            ExcelFilesChazhao excelfileschazhao = new ExcelFilesChazhao();
                            excelfileschazhao.FilesChaozhao(excelinfo);
                            break;
                        case "查找内容位置":
                            ExcelChaZhaoPosition excelchazhaoposition = new ExcelChaZhaoPosition();
                            excelchazhaoposition.ChaZhaoPosition(excelinfo);
                            break;
                        case "行高":
                            ExcelRowHeight excelrowheight = new ExcelRowHeight();
                            excelrowheight.RowHeight(excelinfo);
                            break;
                        case "合并居中单元格":
                            ExcelMergeCell excelmergecell = new ExcelMergeCell();
                            excelmergecell.MergeCell(excelinfo);
                            break;
                        case "字体样式":
                            ExcelFont excelfont = new ExcelFont();
                            excelfont.Font(excelinfo);
                            break;
                        case "字体大小":
                            ExcelFontSize excelfontsize = new ExcelFontSize();
                            excelfontsize.FontSize(excelinfo);
                            break;
                        case "字体格式":
                            ExcelFontStyle excelfontstyle = new ExcelFontStyle();
                            excelfontstyle.FontStyle(excelinfo);
                            break;
                        case "字体颜色":
                            ExcelFontColor excelfontcolor = new ExcelFontColor();
                            excelfontcolor.FontColor(excelinfo);
                            break;
                        case "单元格颜色":
                            ExcelCellColor excelcellcolor = new ExcelCellColor();
                            excelcellcolor.CellColor(excelinfo);
                            break;
                        case "单元格公式":
                            ExcelJudgeFormula exceljudgeformula = new ExcelJudgeFormula();
                            exceljudgeformula.JudgeFormula(excelinfo);
                            break;
                        case "判断筛选":
                            ExcelJudgeFiltrate exceljudgefiltrate = new ExcelJudgeFiltrate();
                            exceljudgefiltrate.JudgeFiltrate(excelinfo);
                            break;
                        //添加列宽
                        case "列宽":
                            ExcelCellWidth excelcellwidth = new ExcelCellWidth();
                            excelcellwidth.CellWidth(excelinfo);
                            break;
                        //添加单元格格式
                        case "单元格格式":
                            ExcelCellStyle excelcellstyle = new ExcelCellStyle();
                            excelcellstyle.CellStyle(excelinfo);
                            break;
                        //添加查找工作表内容位置    
                        case "查找工作表内容位置":
                            ExcelSheetPosition excelsheetposition = new ExcelSheetPosition();
                            excelsheetposition.SheetPosition(excelinfo);
                            break;
                        //添加单元格边框线
                        case "单元格边框线":
                            ExcelCellBorder excelcellborder = new ExcelCellBorder();
                            excelcellborder.CellBorder(excelinfo);
                            break;
                        //添加单元格水平对齐方式
                        case "单元格水平对齐方式":
                            ExcelHorizonAlignment excelhorizonalignment = new ExcelHorizonAlignment();
                            excelhorizonalignment.HorizonAlinment(excelinfo);
                            break;
                        //添加单元格垂直对齐方式
                        case "单元格垂直对齐方式":
                            ExcelVerticalAlignment excelverticalalignment = new ExcelVerticalAlignment();
                            excelverticalalignment.VerticalAlinment(excelinfo);
                            break;
                        default:
                            break;
                    }
            
                }
            }
        }
        #endregion

    }

}