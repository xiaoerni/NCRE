using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DAL;
using Model;
using System.Data;
using System.Data.SqlClient;
//using System.Threading.Tasks;
using System.Collections;
using System.Windows.Forms;

namespace BLL
    
{
    /// <summary>
    /// 得到每个操作题模块的得分情况，  李少然
    /// </summary>
    public  class TotalScoreBLL
    {
       private  TotalScoreDAL totalscoredal;
       public TotalScoreBLL()
        {
                //创建一个winQuestionDAL
            totalscoredal = new DAL.TotalScoreDAL ();
         }

       #region word 根据学号查出 正确答案，学生答案，分数 李少然
       public DataTable ShowWordMesbll(StudentInfoEntity studentinfo)
       {
           return totalscoredal.WordTotalScore(studentinfo);
       } 
       #endregion

       #region ppt 根据学号查出 正确答案，学生答案，分数 李少然
       public DataTable ShowPptMesbll(StudentInfoEntity studentinfo)
       {
           return totalscoredal.PptTotalScore (studentinfo);
       }
       #endregion

       #region windows 根据学号查出 正确答案，学生答案，分数 李少然
       public DataTable ShowWindowsMesbll(StudentInfoEntity studentinfo)
       {
           return totalscoredal.WindowsScore (studentinfo);
       }
       #endregion

       #region IE 根据学号查出 正确答案，学生答案，分数 李少然
       public DataTable ShowIEMesbll(StudentInfoEntity studentinfo)
       {
           return totalscoredal.IETotalScore (studentinfo);
       }
       #endregion

       #region Excel 根据学号查出 正确答案，学生答案，分数 李少然
       public DataTable ShowExcelMesbll(StudentInfoEntity studentinfo)
       {
           return totalscoredal.ExcelTotalScore (studentinfo);
       }
       #endregion


        //  public DataTable LoadWindowsByFlag(winquestion)
        //{
        //    DataTable winQuestionDt = new DataTable();
        //    winQuestionDt = winQuestionDal.LoadWindowsByFlag(winquestion);

        //    int num = winQuestionDt.Rows.Count;      //查询到的datatable的行数

        //    if (num == 0)
        //    {
        //        MessageBox.Show("抽题失败，请联系管理员");

        //    }

      
    }
}
