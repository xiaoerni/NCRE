/*
 * 创建人：赵寒
 * 创建时间：2014-11-10
 * 说明：新建邮件
 * 版权所有：TGB赵寒
 */
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
using System.Data.SqlClient;


namespace NCRE学生考试端V1._0
{
    public partial class frmNewM : Form
    {
        public frmNewM()
        {
            InitializeComponent();
        }

        public static string Context;
        public static string Topic;
        public static string Address;
        public static string BoxPath;
        

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void frmNewM_Load(object sender, EventArgs e)
        {

        }

        private void toolSent_Click(object sender, EventArgs e)
        {

            ///// <summary>
            ///// 查看邮件内容
            ///// </summary>
            ///// <param name="iequestion"></param>

            //IEQuestionEntity iequestion = new IEQuestionEntity();
        
            //IEQuestionEntityBLL iequestionbll = new IEQuestionEntityBLL();
            ////将正确答案，分值取出来，传给studentRecord
          
            //iequestion.questionFlag = "邮件内容";

            //DataTable ieQuestionDt = iequestionbll.LoadIEByFlag(iequestion);

            //IEQuestionRecordEntity   studentRecord = new IEQuestionRecordEntity();

            //studentRecord.studentID = FrmLogin.studentID;
            //string fraction;
            ////string examAnswer;

            //frmNewM frmnewm = new frmNewM();

            Context = txtContent.Text.Trim();
            ////循环遍历正确答案
            //for (int i = 0; i < ieQuestionDt.Rows.Count; i++)
            //{
            //    //将考生ID传到studentRecord实体
            //    studentRecord.studentID = FrmLogin.studentID;
            //    //将试题的ID取出来
            //    studentRecord.questionID = Convert.ToDouble(ieQuestionDt.Rows[i]["questionID"]); 
            //    //将题的分数取出来
            //    fraction = ieQuestionDt.Rows[i]["fraction"].ToString();              
            //    //将考生答案保存
            //    studentRecord.examAnswer = txtContent.Text;
            //    if (txtContent .Text .Trim() == ieQuestionDt.Rows[i]["correctAnswer"].ToString())
            //    {
            //        studentRecord.fraction= Convert.ToDouble(fraction);

            //    }
            //    else
            //    {
            //        studentRecord.fraction= 0;
            //    }
            //    iequestionbll.ReturnScore(studentRecord);
                
                
            // }
       ///// <summary>
       ///// 查看邮件标题
       ///// </summary>
       ///// <param name="iequestion"></param>
       ///// 
       //     //将正确答案，分值取出来，传给studentRecord
            
       //     iequestion.questionFlag = "邮件标题";

            Topic = txtTopic.Text.Trim();
            ////循环遍历正确答案
            //for (int i = 0; i < ieQuestionDt.Rows.Count; i++)
            //{
            //    //将考生ID传到studentRecord实体
            //    studentRecord.studentID = FrmLogin.studentID;
            //    //将试题的ID取出来
            //    studentRecord.questionID = Convert.ToDouble(ieQuestionDt.Rows[i]["questionID"]);
            //    //将题的分数取出来
            //    fraction = ieQuestionDt.Rows[i]["fraction"].ToString();              
            //    //将考生答案保存
            //    studentRecord.examAnswer =txtTopic.Text;
            //    if (txtTopic.Text.Trim() == ieQuestionDt.Rows[i]["correctAnswer"].ToString())
            //    {
            //        studentRecord.fraction = Convert.ToDouble(fraction);
            //    }
            //    else
            //    {
            //        studentRecord.fraction= 0;
            //    }
            //    iequestionbll.ReturnScore(studentRecord);
            //}
       

       /// <summary>
       /// 查看邮件地址
       /// </summary>
       /// <param name="iequestion"></param>
      
            //将正确答案，分值取出来，传给studentRecord
          
            //iequestion.questionFlag = "邮件地址";

            Address = txtAddress.Text.Trim();
          //循环遍历正确答案
            //for (int i = 0; i < ieQuestionDt.Rows.Count; i++)
            //{
            //    ///将考生ID传到studentRecord实体
            //    studentRecord.studentID = FrmLogin.studentID;
            //    //将试题的ID取出来
            //    studentRecord.questionID = Convert.ToDouble(ieQuestionDt.Rows[i]["questionID"]);
            //    //将题的分数取出来
            //    fraction = ieQuestionDt.Rows[i]["fraction"].ToString();              
            //    //将考生答案保存
            //    studentRecord.examAnswer = txtAddress.Text ;
            //    if (txtAddress.Text.Trim() == ieQuestionDt.Rows[i]["correctAnswer"].ToString())
            //    {
            //        studentRecord.fraction= Convert.ToDouble(fraction);
            //    }
            //    else
            //    {
            //        studentRecord.fraction = 0;
            //    }
            //    iequestionbll.ReturnScore(studentRecord);
            //}
            /// <summary>
            /// 查看附件内容
            /// </summary>
            /// <param name="iequestion"></param>

            
            //将正确答案，分值取出来，传给studentRecord
           
            //iequestion.questionFlag = "邮件附件";

             BoxPath = txtboxPath.Text.Trim();
            ////循环遍历正确答案
            //for (int i = 0; i < ieQuestionDt.Rows.Count; i++)
            //{
            //    //将考生ID传到studentRecord实体
            //    studentRecord.studentID = FrmLogin.studentID;
            //    //将试题的ID取出来
            //    studentRecord.questionID = Convert.ToDouble(ieQuestionDt.Rows[i]["questionID"]);
            //    //将题的分数取出来
            //    fraction = ieQuestionDt.Rows[i]["fraction"].ToString();
            //    //将考生答案保存
            //    studentRecord.examAnswer = txtboxPath.Text;
            //    if (txtboxPath.Text.Trim() == ieQuestionDt.Rows[i]["correctAnswer"].ToString())
            //    {
            //        studentRecord.fraction = Convert.ToDouble(fraction);

            //    }
            //    else
            //    {
            //        studentRecord.fraction = 0;
            //    }
            //    iequestionbll.ReturnScore(studentRecord);


            //}
            //判断收件人、主题、内容是否为空
            if (txtAddress.Text.Trim() == "")

                MessageBox.Show("请填写收件人");
            else
            {
                if (txtTopic.Text.Trim() == "")
                {
                    if (MessageBox.Show("确认主题为空？", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                    {
                        MessageBox.Show("邮件发送成功!");
                    }
                }
                else
                     MessageBox.Show("邮件发送成功!");
            }
       }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            ////打开考生目录
            //string path = @"C:\Hb15Dir\" + FrmLogin.studentID;
            //System.Diagnostics.Process.Start("explorer.exe", path);

            OpenFileDialog dialog1 = new OpenFileDialog();
            if (dialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtboxPath.Text = dialog1.SafeFileName;

            }
        }

        private void frmNewM_Load_1(object sender, EventArgs e)
        {
            
        }

        private void lblboxPath_Click(object sender, EventArgs e)
        {
            ////打开考生目录
            //string path = @"C:\Hb15Dir\" + FrmLogin.studentID;
            //System.Diagnostics.Process.Start("explorer.exe", path);

            OpenFileDialog dialog1 = new OpenFileDialog();
            if (dialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtboxPath.Text = dialog1.SafeFileName;

            }
        }

        private void 文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ////打开考生目录
            //string path = @"C:\Hb15Dir\" + FrmLogin.studentID;
            //System.Diagnostics.Process.Start("explorer.exe", path);

            OpenFileDialog dialog1 = new OpenFileDialog();
            if (dialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtboxPath.Text = dialog1.SafeFileName;

            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}

