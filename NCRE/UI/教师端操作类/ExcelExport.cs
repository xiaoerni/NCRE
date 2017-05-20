﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BLL;
using Model;

namespace NCRE学生考试端V1._0
{
    public class ExcelExport
    {
        //private string fileName;
        /// <summary>
        /// 读Excel的数据 在DataGridView中显示
        /// </summary>
        /// <param name="dgv">显示内容的DataGridView的名称</param>
        public void setExcel(DataGridView dgv, string name)
        {
            //总可见列数，总可见行数
            int colCount = dgv.Columns.GetColumnCount(DataGridViewElementStates.Visible);
            int rowCount = dgv.Rows.GetRowCount(DataGridViewElementStates.Visible);
            //dataGridView 没有数据提示
            if (dgv.Rows.Count == 0 || rowCount == 0)
            {
                MessageBox.Show("表中没有数据", "提示");
            }
            else
            {
                //选择创建文件的路径
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "excel files(*.xlsx)|*.xlsx";
                save.Title = "请选择要导出数据的位置";
                save.FileName = name + "--"+DateTime.Now.ToLongDateString()+"--计算机一级考试成绩";
                if (save.ShowDialog() == DialogResult.OK)
                {
                    string fileName = save.FileName;
                    //MessageBox.Show(save.FileName);
                    // 创建Excel对象
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                    if (excel == null)
                    {
                        MessageBox.Show("Excel无法启动", "提示");
                        return;
                    }
                    //创建Excel工作薄
                    Microsoft.Office.Interop.Excel.Workbook excelBook = excel.Workbooks.Add(true);
                    Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets[1];

                    //生成字段名称
                    int k = 0;
                    for (int i = 0; i < dgv.ColumnCount; i++)
                    {
                        if (dgv.Columns[i].Visible)  //不导出隐藏的列
                        {
                            excel.Cells[1, k + 1] = dgv.Columns[i].HeaderText;
                            k++;
                        }
                    }
                    //填充数据
                    for (int i = 0; i < dgv.RowCount; i++)
                    {
                        k = 0;
                        for (int j = 0; j < dgv.ColumnCount; j++)
                        {
                            if (dgv.Columns[j].Visible)  //不导出隐藏的列
                            {
                                if (dgv[j, i].ValueType == typeof(string))
                                {
                                    excel.Cells[i + 2, k + 1] = "" + dgv[j, i].Value.ToString();
                                }
                                else
                                {
                                    excel.Cells[i + 2, k + 1] = dgv[j, i].Value.ToString();
                                }
                            }
                            k++;
                        }
                    }
                    try
                    {
                        excelBook.Saved = true;
                        excelBook.SaveCopyAs(fileName);
                        MessageBox.Show("导出成功！", "提示");
                        
                    }
                    catch
                    {
                        MessageBox.Show("导出失败，文件可能正在使用中", "提示");
                    }

                }
            }
        }
    }

}
