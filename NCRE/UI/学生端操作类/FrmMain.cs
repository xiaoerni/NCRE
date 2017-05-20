using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Model;
using System.Data.SqlClient;
using BLL;
using System.Threading;
using Microsoft.Win32;
using System.Diagnostics;
using NCRE学生考试端V1._0.选择题;
using NCRE学生考试端V1._0.悬浮框;
using System.Runtime.InteropServices;

namespace NCRE学生考试端V1._0
{
    public partial class FrmMain : Form
    {
        public static Form frmmain;
        public FrmMain()
        {
            InitializeComponent();
            st = InitStudent(MyInfo.MystudentID(), MyInfo.MycollegeID());
            secondColumnWidth = this.tableLayoutPanel1.ColumnStyles[1].Width;


            frmmain = this;

        }
        //实例化一个学生实体
        public StudentInfoEntity studentinfo = new StudentInfoEntity();

        /// <summary>
        /// 加载学生信息
        /// </summary>
        public static StudentInfoEntity st = new StudentInfoEntity();
        private StudentInfoEntity InitStudent(string studentId, string collegeId)
        {
            st.studentID = studentId;
            st.CollegeID = collegeId;
            return st;
        }

        public static float firstColumnWidth = 0;
        public static float secondColumnWidth = 0;
        SelectQuestionRecordEntity sqEntity = new SelectQuestionRecordEntity();
        SelectQuestionBLL sqBLL = new SelectQuestionBLL();    //申明一个B层的实例
        int numPage = 0;                                      //记录页数
        List<SelectQuestionRecordEntity> listRecord;       //答题记录实体
        //记录考试倒计时的总时间

        private void FrmMain_Load(object sender, EventArgs e)
        {
            //显示悬浮框
            frmxuanfukuang frmxfk = new frmxuanfukuang();
            frmxfk.Focus();   //让窗体获得焦点
            frmxfk.Show();    //显示窗体
            

            Rectangle ScreenArea = System.Windows.Forms.Screen.GetBounds(this);         
            int height1 = ScreenArea.Height; //屏幕高度

            this.Left = 0;
            this.Top = 0;
            this.Height = height1;
            

            //jiazaixuanzeti
            loadSelectQuestion();

        }
      
        
     

        //页签的切换
        private void tabControl1_Selected_1(object sender, TabControlEventArgs e)
        {

            string tabName = ((System.Windows.Forms.Panel)(((System.Windows.Forms.TabControl)(sender)).SelectedTab)).Text;
            if (tabName == "单项选择")
            {
                tlpSelect.Visible = true;
                this.Width = 100 * Screen.PrimaryScreen.WorkingArea.Width / 100;
                //loadSelectQuestion();
                //this.WindowState = FormWindowState.Minimized;
                tableLayoutPanel1.ColumnStyles[1].Width = secondColumnWidth;
            }
            else
            {
                tlpSelect.Visible = false;
                tableLayoutPanel1.ColumnStyles[1].Width = 0;
                this.Width = 37 * Screen.PrimaryScreen.WorkingArea.Width / 100;
            }
        }

        //上一页
        private void lastPageBtn_Click(object sender, EventArgs e)
        {
            try
            {
                listRecord = sqBLL.GetLstSelectQuestionRecordByStudentIdAndCollegeId(st);
                flowLayoutPanel1.Focus();
                numPage--;    //页数减1
                this.nextPageBtn.Enabled = true;   //下一页可用
                if (numPage <= 1)
                {
                    this.lastPageBtn.Enabled = false;
                }
                flowLayoutPanel1.Controls.Clear();
                //添加用户控件
                for (int i = 0; i < 4; i++)
                {
                    UCSelect uc = new UCSelect();
                   // uc.Dock = DockStyle.Fill;
                    uc.Name = i.ToString();
                    uc.Top = 150 * i + 10;
                    uc.Left = 100;
                    flowLayoutPanel1.Controls.Add(uc);
                    //添加试题
                    int num = i + (numPage - 1) * 4;
                    uc.BindDataToSelf(listRecord[num], num);
                }

                this.lblPageInfo.Text = "第" + numPage + "/" + listRecord.Count / 4 + " 页 ";
            }
            catch (Exception)
            {
                
   
            }

          
        }

        //下一页
        private void nextPageBtn_Click(object sender, EventArgs e)
        {
            try
            {
                listRecord = sqBLL.GetLstSelectQuestionRecordByStudentIdAndCollegeId(st);
                numPage++;    //页数加一          
                //lastPageBtn.Enabled = true;    //上一页可用
                this.lastPageBtn.Enabled = true;
                flowLayoutPanel1.Focus();       //滚动条能动

                if (numPage == 5)
                {
                    //nextPageBtn.Enabled = false;
                    this.nextPageBtn.Enabled = false;
                }

                //添加用户控件
                flowLayoutPanel1.Controls.Clear();
                for (int i = 0; i < 4; i++)
                {
                    UCSelect uc = new UCSelect();
                    uc.Name = i.ToString();
                    uc.Top = 150 * i + 10;
                    uc.Left = 100;
                    flowLayoutPanel1.Controls.Add(uc);
                    //添加试题
                    int num = i + (numPage - 1) * 4;
                    uc.BindDataToSelf(listRecord[num], num);

                }

                this.lblPageInfo.Text = "第" + numPage + "/" + listRecord.Count / 4 + " 页 ";

            }
            catch (Exception)
            {


            }
        }

        #region 交卷的方法HandIn ---孟海滨
        /// <summary>
        /// 交卷的方法
        /// </summary>
        /// <returns>没有做的题目数量</returns>
        public int HandIn()
        {
            //1判断所有选择题是否做完
            try
            {
                listRecord = sqBLL.GetLstSelectQuestionRecordByStudentIdAndCollegeId(st);
                int num = 0;
                for (int i = 0; i < listRecord.Count; i++)
                {
                    if (listRecord[i].ExamAnswer == string.Empty)
                    {
                        num++;
                    }
                }

                return num;
            }
            catch (Exception)
            {

                return 0;
            } 
        
        }
        #endregion

        //交卷
        private void button1_Click(object sender, EventArgs e)
        {
            int num = HandIn();
            if (num > 0)
            {
                DialogResult dr = MessageBox.Show(this, "您还有" + num + "道题还没有做，是否要提交选择题。", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.OK)
                {
                    //this.Close();
                    tlpSelect.Visible = false;
                    tableLayoutPanel1.ColumnStyles[1].Width = 0;
                    this.Width = 37 * Screen.PrimaryScreen.WorkingArea.Width / 100;
                }

            }
            else
            {
                DialogResult dr = MessageBox.Show(this, "是否要提交选择题", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.OK)
                {
                    tlpSelect.Visible = false;
                    tableLayoutPanel1.ColumnStyles[1].Width = 0;
                    this.Width = 37 * Screen.PrimaryScreen.WorkingArea.Width / 100;
                }
            }

            frmMain2 frmmain2 = new frmMain2();
            frmmain2.Show();
            this.Hide();
        }

        private void pannelBtn_Paint(object sender, PaintEventArgs e)
        {

        }

        #region 加载选择题loadSelectQuestion
        /// <summary>
        /// 加载选择题
        /// </summary>
        public void loadSelectQuestion()
        {
            try
            {
                #region 加载试题并切保存答题信息
                //1.查询数据，返回List<QuestionRecordEntity);
                listRecord = sqBLL.GetLstSelectQuestionRecordByStudentIdAndCollegeId(st);
                if (listRecord.Count != 0)
                {
                    //2.将listQuestion添加到用户控件中
                    for (int i = 0; i < 4; i++)
                    {
                        //添加第i个用户控件到窗体上
                        UCSelect uc = new UCSelect();
                        uc.Name = i.ToString();
                        uc.Top = 150 * i + 10;
                        uc.Left = 100;
                        flowLayoutPanel1.Controls.Add(uc);
                        //将试题信息绑定到用户控件上
                        uc.BindDataToSelf(listRecord[i], i + numPage * 4);
                    }

                    //当前页数统计和显示
                    numPage++;
                    if (numPage == 1)
                    {
                        lastPageBtn.Enabled = false;
                    }
                    this.lblPageInfo.Text = "第" + numPage + "/" + listRecord.Count / 4 + " 页 ";
                }
                else
                {
                    MessageBox.Show(this, "您还没有试题，请联系管理员", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }

                #endregion
            }
            catch (Exception)
            {

                throw;
            }
        }
        #endregion

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {

        }



       
       
       
    }
}