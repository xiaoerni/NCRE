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
    public partial class FrmLogin : Form
    {
        public FrmLogin()
        {
            InitializeComponent();
        }     
        public static string studentID;      
        private void btnLogin_Click(object sender, EventArgs e)
        {
          
            //根据学生ID取出学生信息
            StudentInfoEntity studentinfo = new StudentInfoEntity();
            studentinfo.studentID = txtStudentID.Text.Trim();
            StudentInfoEntityBLL studentinfobll = new StudentInfoEntityBLL();
            //List<StudentInfoEntity> stuinfo=  studentinfobll.SelectStudentInfoByID(studentinfo );
            //MessageBox.Show(studentinfo.studentName);
            DataTable dt = studentinfobll.SelectStudentInfoByID(studentinfo);
            //MessageBox.Show(dt.Rows[0]["studentName"].ToString ());
            //在pannel中显示学生信息

            //验证学生是否可以考试
            //StudentInfoBLL studentBll = new StudentInfoBLL();
            //StudentInfoEntity enStudent = studentBll.GetStudentById(studentID);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("学生不存在！");
                return;
            }


            //让登陆键show
            btnOK.Visible = true;
            //让之前的登陆键隐藏
            btnLogin.Visible = false;

            panel1.Visible = true;
            //考场信息
            lblexamPlaceID.Text = dt.Rows[0]["examPlaceID"].ToString();
            //准考证号
            lblstudentID.Text = dt.Rows[0]["studentID"].ToString();
            //学生姓名性别
            lblstuName.Text = dt.Rows[0]["studentName"].ToString();
            lblSex.Text = "（" + dt.Rows[0]["sex"].ToString() + "）";
            //座位号 IP
        }

        private void FrmLogin_Load(object sender, EventArgs e)
        {
            studentID = txtStudentID.Text.Trim();
            panel1.Visible = false;
            btnOK.Visible = false;
        }

      
        private void btnOK_Click(object sender, EventArgs e)
        {
            studentID = txtStudentID.Text.Trim();
            //验证学生是否可以考试
            StudentInfoBLL studentBll = new StudentInfoBLL();
            StudentInfoEntity enStudent = studentBll.GetStudentById(studentID);
            if (enStudent == null)
            {
                MessageBox.Show("学生不存在！");
            }
            UserEntityBLL userDal = new UserEntityBLL();

            if (userDal.GetIsCanExamByStudent(enStudent) == false)
            {
                MessageBox.Show("该学生还未分配考题！");
                this.Close();
            }
            else
            {
                this.Hide();
                FrmMain frmmain = new FrmMain();
                frmmain.ShowDialog();
            }
           

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FrmTeacherLogin frmteacherlogin = new FrmTeacherLogin();
            frmteacherlogin.Show();
        }

        private void txtStudentID_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Hide();
        }


    }
}
