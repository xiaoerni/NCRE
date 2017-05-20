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
    public partial class FrmTeacherLogin : Form
    {
        public FrmTeacherLogin()
        {
            InitializeComponent();
        }

        private void btnReturn_Click(object sender, EventArgs e)
        {
            //返回首页上
            this.Hide();
            //FrmLogin frmlogin = new FrmLogin ();
            //frmlogin.Show();

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            //登陆
            UserEntityBLL userentitybll = new UserEntityBLL();
            UserEntity userinfo = new UserEntity();
            userinfo.userName = txtUsername.Text.Trim();
            userinfo.userPassword = txtPwd.Text.Trim();
            DataTable dt= userentitybll.TeacherLoginByName (userinfo);
            if (dt.Rows.Count  > 0)
            {
                this.Hide();
                frmTeacherManagerMain teacherform = new frmTeacherManagerMain();
                teacherform.Show();

            }
            else
            {
                MessageBox.Show("没有对应的数据，登陆失败");
            }
        }
    }
}
