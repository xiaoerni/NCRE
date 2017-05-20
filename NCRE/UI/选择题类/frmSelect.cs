using NCRE学生考试端V1._0.选择题;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using Model;
using System.Windows.Forms;
using BLL;

namespace NCRE学生考试端V1._0
{
    public partial class frmSelect : Form
    {
        public frmSelect()
        {
            InitializeComponent();
            flowLayoutPanel1.Focus();

            //传递学号和学院ID++++++++++++++++++++++++++++++++++++++=
            st = InitStudent("12080141020", "01");
        }

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
    

        SelectQuestionRecordEntity sqEntity = new SelectQuestionRecordEntity();
        //申明一个B层的实例
        SelectQuestionBLL sqBLL = new SelectQuestionBLL();
        int numPage = 0; //记录页数
        List<SelectQuestionRecordEntity> listRecord;
        private void frmSelect_Load(object sender, EventArgs e)
        {
            
            flowLayoutPanel1.Focus();
            //StudentInfoEntity st = new StudentInfoEntity();
            //#region TODO 添加学生的学号和学院号
            //st.studentID = "11111111";
            //st.collegeID = "01"; 
            //#endregion

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
                    lblPageInfo.Text = "第" + numPage + "/" + listRecord.Count / 4 + " 页 ";
                }
                else
                {
                    MessageBox.Show(this, "您还没有试题，请联系管理员", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }

                #endregion
            }
            catch (Exception  )
            {
                
                throw ;
            }
           
        }

        #region 个不重复随机数GetListRandom
        /// <summary>
        /// 取count个不重复随机数
        /// </summary>
        /// <param name="count">取的个数</param>
        /// <param name="Max_Count">最大范围</param>
        /// <returns>随机数的list集合</returns>
        private List<int> GetListRandom(int count, int Max_Count)
        {
            Random random = new Random();
            int intRan = 0;
            List<int> listRandom = new List<int>();   //定义一个集合，用来存储生成的随机数

            //生成count个随机不相同的数
            for (int i = 0; i < count; i++)
            {
                intRan = Convert.ToInt32(random.Next(1, Max_Count));

                if (listRandom.Contains(intRan))
                {
                    i--;
                }
                else
                {
                    listRandom.Add(intRan);
                }

            }

            return listRandom;    //利用三目运算确定区间的开始位置
        }
        #endregion

        //上一页
        private void lastPageBtn_Click(object sender, EventArgs e)
        {
            try
            {
                listRecord = sqBLL.GetLstSelectQuestionRecordByStudentIdAndCollegeId(st);
                //sqBLL.UpdateSelectQuestionRecordByStudentInfo(st,
                flowLayoutPanel1.Focus();
                numPage--;    //页数减1
                nextPageBtn.Enabled = true;   //下一页可用
                if (numPage <= 1)
                {
                    lastPageBtn.Enabled = false;
                }
                flowLayoutPanel1.Controls.Clear();
                //添加用户控件
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

                lblPageInfo.Text = "第" + numPage + "/" + listRecord.Count / 4 + " 页 ";
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
                lastPageBtn.Enabled = true;    //上一页可用
                flowLayoutPanel1.Focus();       //滚动条能动
                if (numPage == 5)
                {
                    nextPageBtn.Enabled = false;
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

                lblPageInfo.Text = "第" + numPage + "/" + listRecord.Count / 4 + " 页 ";
     
            }
            catch (Exception)
            {
                

            }
               
        }

        //交卷
        private void btnOK_Click(object sender, EventArgs e)
        {
            int num = HandIn();
            if (num > 0)
            {
                DialogResult dr = MessageBox.Show(this, "您还有" + num + "道题还没有做，是否要提交选择题", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.OK)
                {
                    this.Close();
                    
                }

            }
            else
            {
                DialogResult dr = MessageBox.Show(this, "是否要提交选择题", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.OK)
                {
                    this.Close();
                }
            }


        }


        private void frmSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        /// <summary>
        /// 交卷的方法
        /// </summary>
        /// <returns>没有做的题目数量</returns>
        public int  HandIn() {
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

            //1.1查询学生答题记录，返回List<QuestionRecordEntity);
           
          
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        
    }
}
