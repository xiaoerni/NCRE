using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Model;
using BLL;

namespace NCRE学生考试端V1._0.选择题
{
    public partial class UCSelect : UserControl
    {
        SelectQuestionBLL sqBLL = new SelectQuestionBLL();
        string rightAnswer = string.Empty;

        public UCSelect()
        {
            InitializeComponent();
            radioButton5.Checked = true;
        }

        //GetQuestion gq = new GetQuestion();
        private string value = string.Empty;
        public string Value
        {
            get { return this.value; }
            set
            {
                this.value = value;

            }
        }

        /// <summary>
        /// 将信息加载到控件上
        /// </summary>
        /// <param name="qe"></param>
        /// <param name="count"></param>
        public void BindDataToSelf(SelectQuestionRecordEntity qe, int count)
        {
            try
            {
                rightAnswer = qe.RightAnswer;
                label1.Text = qe.QuestionContent.Replace("\r\n", "<br>");
                label2.Text = qe.OptionA.Replace("\r\n", "<br>");
                label3.Text = qe.OptionB.Replace("\r\n", "<br>");
                label4.Text = qe.OptionC.Replace("\r\n", "<br>");
                label5.Text = qe.OptionD.Replace("\r\n", "<br>");
                int num = 0;
                Choice1.Name = qe.QuestionID;

                if (qe.ExamAnswer != string.Empty)
                {
                    foreach (Control item in Choice1.Controls)
                    {
                        if (item is RadioButton && qe.ExamAnswer == item.Text)
                        {
                            RadioButton rb = (RadioButton)item;
                            rb.Checked = true;
                        }
                    }

                }
                num = count + 1;
                Choice1.Text = "选择题" + num;
            }
            catch (Exception)
            {
                

            }
            
         

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.value = radioButton1.Text;
                int numFlag = UpdateSelectQuestionRecord(value);
                if (numFlag == 0)
                {
                    MessageBox.Show(this, "信息保存失败，请联系管理员", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }
            }
            catch (Exception)
            {
                

            }
           
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.value = radioButton2.Text;
                int numFlag = UpdateSelectQuestionRecord(value);
                if (numFlag == 0)
                {
                    MessageBox.Show(this, "信息保存失败，请联系管理员", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }
            }
            catch (Exception)
            {
                

            }
         
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.value = radioButton3.Text;
                int numFlag = UpdateSelectQuestionRecord(value);
                if (numFlag == 0)
                {
                    MessageBox.Show(this, "信息保存失败，请联系管理员", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }
            }
            catch (Exception)
            {

   
            }
            
            
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.value = radioButton4.Text;
                int numFlag = UpdateSelectQuestionRecord(value);
                if (numFlag == 0)
                {
                    MessageBox.Show(this, "信息保存失败，请联系管理员", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }
            }
            catch (Exception)
            {


            }
          
        }

        #region 保存答题记录
        /// <summary>
        /// 保存答题记录
        /// </summary>
        /// <param name="strAnswer"></param>
        /// <returns></returns>
        public int UpdateSelectQuestionRecord(string strAnswer)
        {
            try
            {
                //学生答题信息保存
                StudentInfoEntity st = FrmMain.st;

                SelectQuestionRecordEntity sqRecord = new SelectQuestionRecordEntity();

                sqRecord.QuestionID = Choice1.Name;
                sqRecord.ExamAnswer = strAnswer;
                int numFlag = sqBLL.UpdateSelectQuestionRecordByStudentInfo(st, sqRecord,rightAnswer);
                if (numFlag == 0)
                {
                    MessageBox.Show(this, "111系统错误，请联系管理员", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return 0;
                }
                else
                {
                    return 1;
                }
            }
            catch (Exception)
            {
                return 0;

            }
           

        } 
        #endregion

        private void Choice1_Enter(object sender, EventArgs e)
        {
            
        }
    }
}
