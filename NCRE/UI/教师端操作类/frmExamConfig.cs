using BLL;
using Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NCRE学生考试端V1._0
{
    public partial class frmExamConfig : Form
    {
        public frmExamConfig()
        {
            InitializeComponent();
        }

        #region 共有属性
        ExamConfigBLL examConfigBll = new ExamConfigBLL();
        private List<CollegeEntity> allCollege = new List<CollegeEntity>();

        /// <summary>
        /// 所有的学生 集合
        /// </summary>
        private List<StudentInfoEntity> lstAllStudent = new List<StudentInfoEntity>();
        
        /// <summary>
        /// 选中学生的集合
        /// </summary>
        private List<StudentInfoEntity> lstSelectStudent = new List<StudentInfoEntity>();
        #endregion


        /// <summary>
        /// 窗体加载，加载学院列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmExamConfig_Load(object sender, EventArgs e)
        {
            List<CollegeEntity> lstCollege = examConfigBll.GetAllCollege();
            dgvStudent.AutoGenerateColumns = false;
            cbCollege.Items.Add("全部");
            //加载所有的列表信息
            foreach (CollegeEntity college in lstCollege)
            {
                allCollege.Add(college);
                cbCollege.Items.Add(college.collegeName);
            }
            cbCollege.SelectedIndex = 0;
        }

        /// <summary>
        /// 条件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbCollege_SelectedIndexChanged(object sender, EventArgs e)
        {
            //获取下拉框中的学院名称
            string collegeName = ((System.Windows.Forms.ComboBox)(sender)).SelectedItem.ToString();

            //根据学院名称查询  学院实体
            CollegeEntity enCollege = allCollege.Find(s => s.collegeName == collegeName);
            if (enCollege==null)
            {
                return;
            }
            //查询该学院内的学生列表
            lstAllStudent = examConfigBll.GetStudentByCollege(enCollege);

            dgvStudent.DataSource = lstAllStudent;
        }

        /// <summary>
        /// 情况 选择条件的  选择题 答题记录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClear_Click(object sender, EventArgs e)
        {
            //按照学院 清空 该学院的  选择题 答题记录
            CollegeEntity enCollege = new CollegeEntity();
            if (cbCollege.Text=="全部")
            {
                //如果选中全部，则清空所有的选择题答题记录
                examConfigBll.ClearSelectQuestionRecordByCollegeID(allCollege);
            }

            //在全集中查找选中的学院
            enCollege= allCollege.Find(s => s.collegeName == cbCollege.Text);
            if (enCollege==null)
            {
                return;
            }

            if (lstSelectStudent.Count == 0)
            {
                //清空指定学院内的选择题答题记录
                examConfigBll.FalseClearSelectQuestionRecordByCollegeID(enCollege);
            }
            else { 
                //清空选中学生的选择题答题记录
                examConfigBll.DeleteSelectQuestionRecordByLstStudent(lstSelectStudent);
            }

            MessageBox.Show("该学院的选择题记录清空成功！");
        }

        /// <summary>
        /// 对选中学院内的所有学生进行抽题，耗时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfigRecord_Click(object sender, EventArgs e)
        {
            //注意：耗时操作。根据所选学生，进行抽题。然后写入到答题记录表中
            //对所选学生，进行抽题（选择题）
            //TSMCIS.ClientFramework.Development.Utility.ShowLoading(this);
            if (lstSelectStudent.Count == 0)
            {
                //如果选中的学生为空，则默认整个学院的学生进行配置考试
                examConfigBll.RandGenerateRecord(lstAllStudent);
            }
            else {
                //如果有选中学生，则对选中的学生进行选择题抽题
                examConfigBll.RandGenerateRecord(lstSelectStudent);
            }
            //TSMCIS.ClientFramework.Development.Utility.HideLoading();

            MessageBox.Show("该学院的选择题答题记录配置成功！");
        }


        /// <summary>
        /// 选中学生
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvStudent_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //先查询学生
            StudentInfoEntity seletStudent = new StudentInfoEntity();
            seletStudent = lstAllStudent.Find(s => s.studentID == dgvStudent.Rows[e.RowIndex].Cells["studentID"].Value.ToString());
            //学生不存在，则退出
            if (seletStudent == null)
            {
                return;
            }

            if (dgvStudent.Rows[e.RowIndex].Cells["isCheck"].Value == null)
            {
                dgvStudent.Rows[e.RowIndex].Cells["isCheck"].Value = true;
                //初次选中
                lstSelectStudent.Add(seletStudent);
                return;
            }

            //选中的学生存在
            if (false == (bool)dgvStudent.Rows[e.RowIndex].Cells["isCheck"].Value)
            {
                //如果之前是未选中状态，则进行选中
                lstSelectStudent.Add(seletStudent);
                dgvStudent.Rows[e.RowIndex].Cells["isCheck"].Value = true;
            }
            else { 
                //之前是选中状态
                lstSelectStudent.Find(s => s.studentID == seletStudent.studentID);
                if (lstSelectStudent==null)
                {
                    return;
                }
                //之前处于选中状态
                lstSelectStudent.Remove(seletStudent);
                dgvStudent.Rows[e.RowIndex].Cells["isCheck"].Value = false;                
            }
        }
    }
}
