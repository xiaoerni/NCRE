using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BLL;
using System.Data;
using Model;

namespace NCRE学生考试端V1._0
{
    public class ExcelLoading
    {
        public ExcelEntityBLL excelbll = new ExcelEntityBLL();

        public DataTable LoadQuestionContent(ExcelQuestionEntity  excelinfo)
        {
            //excelinfo.StudentID = MyInfo.MystudentID();
            return excelbll.LoadExcelQuestion(excelinfo);                            
        }
    }
}
