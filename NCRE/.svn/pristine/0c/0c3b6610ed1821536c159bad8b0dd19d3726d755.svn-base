using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;
using BLL;
using System.IO;
using System.Data;
using System.Data.SqlClient;

namespace NCRE学生考试端V1._0
{
    public class WindowsQuestionFlag
    {
        private WinQuestionEntityBLL winquestionbll = new WinQuestionEntityBLL();
        
        private WinQuestionEntity winquestion = new WinQuestionEntity();
        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();

        #region"判断是什么类型的题-韩梦甜-2015-11-20"
        /// <summary>
        /// 判断是什么类型的题
        /// </summary>
        /// <param name="winquestion"></param>
        public void SwitchQuestionFlag(WinQuestionEntity winquestion)
        {
            DataTable dt = new DataTable();  
            dt = winquestionbll.LoadWindowsQuestion(winquestion);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string questionflag = dt.Rows[i]["QuestionFlag"].ToString().Trim();
                String collegeId = MyInfo.MycollegeID();
                if (examconfigbll.IsTableExist("WinQuestionRecordEntity_" + collegeId) == false)
                {
                    break;
                }
                else
                {

                    switch (questionflag)
                    {
                        case "查找文件夹":

                            WindowsFindDirectory windowsfinddirectory = new WindowsFindDirectory();
                            windowsfinddirectory.FindDirectory(winquestion);
                            break;

                        case "查找文件":
                            WindowsFindFile windowsfindfile = new WindowsFindFile();
                            windowsfindfile.FindFile(winquestion);
                            break;

                        case "查找快捷方式":
                            WindowsFindIWshShortcut windowsfindiwshshortcut = new WindowsFindIWshShortcut();
                            windowsfindiwshshortcut.FindIWshShortcut(winquestion);
                            break;
                        case "删除文件夹":
                            WindowsDeleteDirectory windowsdeletedirectory = new WindowsDeleteDirectory();
                            windowsdeletedirectory.DeleteDirectory(winquestion);
                            break;
                        case "只读":
                            WindowsReadOnly windowsreadonly = new WindowsReadOnly();
                            windowsreadonly.ReadOnly(winquestion);
                            break;
                        case "隐藏":
                            WindowsHidden windowshidden = new WindowsHidden();
                            windowshidden.Hidden(winquestion);
                            break;

                        case "开头查找文件":
                        WindowsFindStartFiles windowsfindstart = new WindowsFindStartFiles();
                        windowsfindstart.FindStartFiles(winquestion);
                        break;

                        case "后缀名查找文件":
                        WindowsFindExpendFiles windowsfindexpend = new WindowsFindExpendFiles();
                        windowsfindexpend.FindExpendFiles(winquestion);
                        break;
                }

                    }
                }
            }
        }
#endregion
    }
