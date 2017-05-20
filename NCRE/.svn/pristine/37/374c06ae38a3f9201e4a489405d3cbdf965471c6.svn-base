using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace NCRE学生考试端V1._0.悬浮框
{
    public partial class frmxuanfukuang : Form
    {
        //李芬
        public static Form frmxuanfuk;
        public frmxuanfukuang()
        {
            frmxuanfuk = this;
            InitializeComponent();
        }

        private void xianshibutton2_Click(object sender, EventArgs e)
        {


            frmMain2.frmmain2.Show();

            //xianshibutton2.Enabled = false;
            xianshibutton2.Visible = false;
            hidebutton1.Visible = true;

            //hidebutton1.Enabled = true;
            //    FrmMain frmmain = FrmMain.InstanceObject();	//实例化窗体
            //    frmmain.Focus();   //让窗体获得焦点
            //    frmmain.Show();    //显示窗体
        }
        public void showbtn() {
            frmxuanfuk.Show();
            xianshibutton2.Visible = true;
            hidebutton1.Visible = false;
            
        }
        private void frmxuanfukuang_Load(object sender, EventArgs e)
        {
          

            timer1.Enabled = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            //this.Opacity = 0.5;
            this.Top = 0;
            this.Left = Screen.PrimaryScreen.Bounds.Width - 860;
            this.Width = 505;
            this.Height = 55;
            xianshibutton2.Visible = false;

            //hidebutton1.Visible = false;
            //button1.Visible = false;
            //jiaojuanbutton3.Visible = false;

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            int second = 0;//初始化变量为0
            second = int.Parse(minuestext.Text) * 60 + int.Parse(secondtext.Text);
            //计算文本框中一共是多少秒

            if (second > 0)
            {
                second = int.Parse(minuestext.Text) * 60 + int.Parse(secondtext.Text);
                second = second - 1;

                minuestext.Text = (second / 60).ToString();
                //符号/表示取整数
                secondtext.Text = (second % 60).ToString();
                //符号%表示取余数
                return;
            }
            else
            {
                timer1.Enabled = false;
                MessageBox.Show("考试时间到,请马上交卷");
                this.Close();
                frmxuanfuk.Close();
                FrmJudge juge = new FrmJudge();
                juge.Show ();
            }

        }

        private void hidebutton1_Click(object sender, EventArgs e)
        {
            //FrmMain.frmmain.Hide();
            frmMain2.frmmain2.Hide();
            xianshibutton2.Visible = true;
            //xianshibutton2.Enabled = true;
            //hidebutton1.Enabled = false;
            hidebutton1.Visible = false;
        }

        private void jiaojuanbutton3_Click(object sender, EventArgs e)
        {
            if (frmMain2.frmmain2==null )
            {
                MessageBox.Show("先提交选择题，再点击交卷！");
            }
            else
	{
            frmMain2.frmmain2.Hide();
            FrmJudge frmjudge = new FrmJudge();
            frmjudge.ShowDialog();
           frmxuanfuk.Close();
            
	}
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //打开考生目录
            string path = @"D:\计算机一级考生文件\";
            System.Diagnostics.Process.Start("explorer.exe", path);
        }
    }
}
