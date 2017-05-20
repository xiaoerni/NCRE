using System;
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
    public partial class frmShowMessage : Form
    {
        public frmShowMessage()
        {
            InitializeComponent();
        }
        #region 得到Windows部分，学生的得分情况 李少然
        /// <summary>
        /// 得到Windows部分，学生的得分情况 李少然
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWindows_Click(object sender, EventArgs e)
        {
            //得到学生学号
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = FrmLogin.studentID;
            //得到查询到的数据]
            DataTable dt = new DataTable();
            TotalScoreBLL totalbll = new TotalScoreBLL();
            //将数据添加到datagridview表中
            dt = totalbll.ShowWindowsMesbll(studentinfo);
            dataGridView1.DataSource = null;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.DataSource = dt;
        } 
        #endregion


        #region 得到word部分 学生的得分情况 李少然
        /// <summary>
        /// 得到Word部分学生的得分情况
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnShowWord_Click(object sender, EventArgs e)
        {
            //得到学生学号
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = FrmLogin.studentID;
            //得到查询到的数据]
            DataTable dt = new DataTable();
            TotalScoreBLL totalbll = new TotalScoreBLL();
            //将数据添加到datagridview表中
            dt = totalbll.ShowWordMesbll(studentinfo);
            dataGridView1.DataSource = null;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.DataSource = dt;

        } 
        #endregion

        private void frmShowMessage_Load(object sender, EventArgs e)
        {

        }

        #region 得到EXCEL部分 学生的得分情况 李少然
        /// <summary>
        /// 得到EXCEL部分 学生的得分情况
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExcel_Click(object sender, EventArgs e)
        {
            //得到学生学号
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = FrmLogin.studentID;
            //得到查询到的数据]
            DataTable dt = new DataTable();
            TotalScoreBLL totalbll = new TotalScoreBLL();
            //将数据添加到datagridview表中
            dt = totalbll.ShowExcelMesbll(studentinfo);
            dataGridView1.DataSource = null;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.DataSource = dt;
        } 
        #endregion

        #region 得到PPT部分，学生的得分情况 李少然
        /// <summary>
        /// 得到PPT部分，学生的得分情况 李少然
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPPT_Click(object sender, EventArgs e)
        {
            //得到学生学号
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = FrmLogin.studentID;
            //得到查询到的数据]
            DataTable dt = new DataTable();
            TotalScoreBLL totalbll = new TotalScoreBLL();
            //将数据添加到datagridview表中
            dt = totalbll.ShowPptMesbll(studentinfo);
            dataGridView1.DataSource = null;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.DataSource = dt;
        } 
        #endregion

        #region 得到IE部分，学生的得分情况 李少然
        /// <summary>
        /// 得到IE部分，学生的得分情况 李少然
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnIE_Click(object sender, EventArgs e)
        {
            //得到学生学号
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = FrmLogin.studentID;
            //得到查询到的数据]
            DataTable dt = new DataTable();
            TotalScoreBLL totalbll = new TotalScoreBLL();
            //将数据添加到datagridview表中
            dt = totalbll.ShowIEMesbll(studentinfo);
            dataGridView1.DataSource = null;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.DataSource = dt;
        } 
        #endregion
    }
}
