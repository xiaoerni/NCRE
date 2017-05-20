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
    public partial class frmTeacherManagerMain : Form
    {
        public frmTeacherManagerMain()
        {
            InitializeComponent();
        }
        //定义一个学生分数的逻辑类
        StudentScoreBLL studentscorebll = new StudentScoreBLL();

        //定义一个学生的实体用来传递学生信息
        StudentInfoEntity studentinfo = new StudentInfoEntity();

        //定义一个ExamConfigBLL类
        ExamConfigBLL examconfigbll = new ExamConfigBLL();

        //定义一个WordQuestionEntityBLL类
        WordQuestionEntityBLL wordquestionbll = new WordQuestionEntityBLL();

        //定义一个WinQuestionEntityBLL类
        WinQuestionEntityBLL winquestionbll = new WinQuestionEntityBLL();

        //定义一个ExcelQuestionEntityBLL类
        ExcelEntityBLL excelQuestionBll = new ExcelEntityBLL();

        //定义一个IEBLL类
        IEQuestionEntityBLL iequestionbll = new IEQuestionEntityBLL();

        //定义一个StudentBindPaperTypeBLL类
        StudentBindPaperTypeBLL studentBindPaperType = new StudentBindPaperTypeBLL();

        //定义一个PPTBLL类
        PptQuestionEntityBLL pptQuestionBll = new PptQuestionEntityBLL();

        //定义一个学生类
        StudentInfoBLL studentinfoBll = new StudentInfoBLL();

        /// <summary>
        /// 绑定窗体加载时需要的数据--周洲--2015年11月16日
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmTeacherManagerMain_Load(object sender, EventArgs e)
        {

            //定义一个DataTable为学院选择绑定数据
            DataTable dtcollege = new DataTable();
            //调用方法选择所有的专业
            dtcollege = studentscorebll.SelectAllCollege();
            //为专业的combobox控件赋值
            cboCollege.DataSource = dtcollege;
            cboCollege.DisplayMember = "学院";
            cboCollege.ValueMember = "collegeName";

            //将下拉框中的值传进studentinfo中
            studentinfo.major = "汉语言文学";
            dgvScore.DataSource = studentscorebll.SelectScoreByMajor(studentinfo);
 
        }

        private void dgvScore_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        /// <summary>
        /// 在选择专业的时候选择该专业下的所有学生
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboMajor_SelectedIndexChanged(object sender, EventArgs e)
        {
            //将下拉框中的值传进studentinfo中
            studentinfo.major = cboMajor.Text;
            dgvScore.DataSource = studentscorebll.SelectScoreByMajor(studentinfo);
        }

        private void cboCollege_SelectedIndexChanged(object sender, EventArgs e)
        {        
            //定义一个DataTable为专业选择绑定数据
            DataTable dtmajor = new DataTable();
            switch (cboCollege.Text)
            {
                case "文学院":
                    studentinfo.CollegeID = "01";
                    break;
                case  "社会发展学院":
                    studentinfo.CollegeID = "02";
                    break;
                case "外国语学院":
                    studentinfo.CollegeID = "03";
                    break;
                case "管理学院":
                    studentinfo.CollegeID = "04";
                    break;
                case "美术学院":
                    studentinfo.CollegeID = "05";
                    break;
                case "音乐学院":
                    studentinfo.CollegeID = "06";
                    break;
                case "教育学院":
                    studentinfo.CollegeID = "07";
                    break;
                case "数学与信息科学学院":
                    studentinfo.CollegeID = "08";
                    break;
                case "物理与电子信息学院":
                    studentinfo.CollegeID = "09";
                    break;
                case "化学与材料科学学院":
                    studentinfo.CollegeID = "10";
                    break;
                case "生命科学学院":
                    studentinfo.CollegeID = "11";
                    break;
                case "经济学院":
                    studentinfo.CollegeID = "12";
                    break;
                case "建筑工程学院":
                    studentinfo.CollegeID = "13";
                    break;
                case "体育学院":
                    studentinfo.CollegeID = "14";
                    break;
                case "社科部":
                    studentinfo.CollegeID = "15";
                    break;
                case "其他":
                    studentinfo.CollegeID = "16";
                    break;
                default:
                    studentinfo.CollegeID = "01";
                    break;
            }

            //调用方法选择所有的专业
            dtmajor = studentscorebll.SelectMajorByCollegeID(studentinfo);
            //为专业的combobox控件赋值
            cboMajor.DataSource = dtmajor;
            cboMajor.DisplayMember = "专业";
            cboMajor.ValueMember = "major";
        }

        private void btnCalssScore_Click(object sender, EventArgs e)
        {
            ExcelExport excelexcport = new ExcelExport();
            for (int i = 0; i < cboMajor.Items.Count; i++)
            {
                cboMajor.DataSource = studentscorebll.SelectAllMajor();
                studentinfo.major = cboMajor.GetItemText(cboMajor.Items[i]);
                dgvScore.DataSource = studentscorebll.SelectScoreByMajor(studentinfo);
                excelexcport.setExcel(dgvScore, cboMajor.GetItemText(cboMajor.Items[i]));
            }

        }

        private void btnCollegeScore_Click(object sender, EventArgs e)
        {
            ExcelExport excelexcport = new ExcelExport();
            for (int i = 0; i < cboCollege.Items.Count; i++)
            {
                cboMajor.DataSource = studentscorebll.SelectAllCollege();

                switch (cboCollege.GetItemText(cboCollege.Items[i]))
                {
                    case "文学院":
                        studentinfo.CollegeID = "01";
                        break;
                    case "社会发展学院":
                        studentinfo.CollegeID = "02";
                        break;
                    case "外国语学院":
                        studentinfo.CollegeID = "03";
                        break;
                    case "管理学院":
                        studentinfo.CollegeID = "04";
                        break;
                    case "美术学院":
                        studentinfo.CollegeID = "05";
                        break;
                    case "音乐学院":
                        studentinfo.CollegeID = "06";
                        break;
                    case "教育学院":
                        studentinfo.CollegeID = "07";
                        break;
                    case "数学与信息科学学院":
                        studentinfo.CollegeID = "08";
                        break;
                    case "物理与电子信息学院":
                        studentinfo.CollegeID = "09";
                        break;
                    case "化学与材料科学学院":
                        studentinfo.CollegeID = "10";
                        break;
                    case "生命科学学院":
                        studentinfo.CollegeID = "11";
                        break;
                    case "经济学院":
                        studentinfo.CollegeID = "12";
                        break;
                    case "建筑工程学院":
                        studentinfo.CollegeID = "13";
                        break;
                    case "体育学院":
                        studentinfo.CollegeID = "14";
                        break;
                    case "社科部":
                        studentinfo.CollegeID = "15";
                        break;
                    case "其他":
                        studentinfo.CollegeID = "16";
                        break;
                    default:
                        studentinfo.CollegeID = "01";
                        break;
                }

                dgvScore.DataSource = studentscorebll.SelectScoreByCollege(studentinfo);
                excelexcport.setExcel(dgvScore, cboCollege.GetItemText(cboCollege.Items[i]));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //frmExamConfig frmexamconfig = new frmExamConfig();
            //frmexamconfig.Show();
            //this.WindowState = FormWindowState.Minimized;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("正在配题，可能需要十几秒的时间，耐心等候......");
             //1、从College中查找出所有的学生，这里存一个CollegeID,用到的方法是：
            //定义一个DataTable为学院选择绑定数据
            DataTable dtcollege = new DataTable();
            //调用方法选择所有的专业
            dtcollege = studentscorebll.SelectAllCollegeInfo();
            //2、从College中查找所有的学生。
            CollegeEntity college = new CollegeEntity();
            String collegeId="";


   
            //++++++++++++++++++++++++Word+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            //System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            //stopwatch.Start(); //  开始监视代码运行时间

            //3、WordQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable worddt = new DataTable();
            worddt = wordquestionbll.WordPaperType();

            List<List<WordQuestionEntity>> wordquestionlist = new List<List<WordQuestionEntity>>();
            String Papertype = "";
            if (worddt.Rows.Count != 0)
            {
                for (int i = 0; i < worddt.Rows.Count; i++)
                {
                    Papertype = worddt.Rows[i]["PaperType"].ToString();
                    wordquestionlist.Add(wordquestionbll.WordPaperTypeGroupByPaperType(Papertype));
                }

                //4、给学生分配数据的PaperType,添加WordRecord表             
                foreach (DataRow dr in dtcollege.Rows)
                {
                    collegeId = dr["collegeID"].ToString();
                    college.collegeID = collegeId;
                    if (examconfigbll.IsTableExist("WordQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("WordQuestionRecordEntity_" + collegeId, "WordQuestionRecordEntity");
                    }

                    List<StudentBindPaperTypeEntity> studentBindPaperTypeEntityList = new List<StudentBindPaperTypeEntity>();      //这里在StudentBindPaperTypeEntity中insert
                    List<StudentInfoEntity> studentinfolist = examconfigbll.GetStudentByCollege(college);



                    if (studentinfolist == null)
                    {
                        continue;
                    }

                    if (studentinfolist.Count > 0)
                    {
                        //进行分配
                        for (int j = 0; j < studentinfolist.Count; j++)
                        {
                            List<WordQuestionRecordEntity> wordrecoredStudentlist = new List<WordQuestionRecordEntity>();

                            String strStudentID = studentinfolist[j].studentID;
                            WordQuestionRecordEntity wordStudent = new WordQuestionRecordEntity();
                            wordStudent.StudentID = strStudentID;
                            if (wordquestionbll.SelectWordRecord(wordStudent) == true)
                            {
                                int i = wordquestionlist.Count;
                                int d = 0;
                                d = j % i;
                                for (int m = 0; m < wordquestionlist[d].Count; m++)
                                {
                                    WordQuestionRecordEntity wordrecoredStudent = new WordQuestionRecordEntity();
                                    wordrecoredStudent.StudentID = strStudentID;
                                    wordrecoredStudent.QuestionID = wordquestionlist[d][m].QuestionID;
                                    wordrecoredStudent.QuestionContent = wordquestionlist[d][m].QuestionContent;
                                    wordrecoredStudent.PaperType = wordquestionlist[d][m].PaperType;
                                    wordrecoredStudent.RightAnswer = wordquestionlist[d][m].RightAnswer;
                                    wordrecoredStudentlist.Add(wordrecoredStudent);
                                }
                                wordquestionbll.InsertWordRecordList(wordrecoredStudentlist);


                                StudentBindPaperTypeEntity studentBindPaperTypeEntity = new StudentBindPaperTypeEntity();
                                studentBindPaperTypeEntity.StudentID = studentinfolist[j].studentID;
                                studentBindPaperTypeEntity.CollegeID = studentinfolist[j].CollegeID;
                                studentBindPaperTypeEntity.PaperType = wordquestionlist[d][0].PaperType;
                                studentBindPaperTypeEntity.IsUse =1;
                                if (studentBindPaperType.SelectRecord(studentBindPaperTypeEntity) == true)
                                {
                                    studentBindPaperTypeEntityList.Add(studentBindPaperTypeEntity);
                                }
                            }                           
                        }
                        //-------------------------选择题配题-----------------------------------------------
                        //如果选中全部，则清空所有的选择题答题记录
                        //examConfigBll.ClearSelectQuestionRecordByCollegeID(allCollege);
                        examconfigbll.ClearSelectQuestionRecordByLstStudent(studentinfolist);
                        examconfigbll.RandGenerateRecord(studentinfolist);
                        studentBindPaperType.InsertRecordList(studentBindPaperTypeEntityList);
                    }
                }
            }




            //++++++++++++++++++++++++Win+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


            //WinQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable winddt = new DataTable();
            winddt = winquestionbll.WinPaperType();

            List<List<WinQuestionEntity>> winquestionWhich = new List<List<WinQuestionEntity>>();

            String WinPapertype = "";

            if (winddt.Rows.Count != 0)
            { 
                                     
                for (int i = 0; i < winddt.Rows.Count; i++)
                {
                    WinPapertype = winddt.Rows[i]["paperType"].ToString();
                    winquestionWhich.Add(winquestionbll.WinPaperTypeGroupByPaperType(WinPapertype));
                }

                           
                foreach (DataRow dr in dtcollege.Rows)
                {
                    

                    collegeId= dr["collegeID"].ToString();
                    college.collegeID =collegeId;
                    if (examconfigbll.IsTableExist("WinQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("WinQuestionRecordEntity_" + collegeId, "WinQuestionRecordEntity");
                    }
                    List<StudentInfoEntity> studentinfolist = examconfigbll.GetStudentByCollege(college);

                    if (studentinfolist == null) 
                    { 
                        continue;
                    }

                    if (studentinfolist.Count>0)
                    {
                  
                        //进行分配
                        for (int j = 0; j < studentinfolist.Count; j++)
                        {
                            //做循环
                            List<WinQuestionRecordEntity> winrecoredStudentlist = new List<WinQuestionRecordEntity>();

                            String StrStudentID = studentinfolist[j].studentID;
                            WinQuestionRecordEntity winStudent = new WinQuestionRecordEntity();
                            winStudent.studentID = StrStudentID;
                            if (winquestionbll.SelectWinRecord(winStudent) == true)
                            {
                                int i = winquestionWhich.Count;
                                int d = 0;
                                d = j % i;
                                for (int m = 0; m < winquestionWhich[d].Count; m++)
                                {
                                    WinQuestionRecordEntity winrecoredStudent = new WinQuestionRecordEntity();
                                    winrecoredStudent.studentID = StrStudentID;
                                    winrecoredStudent.questionID = winquestionWhich[d][m].questionID;
                                    winrecoredStudent.questionContent = winquestionWhich[d][m].questionContent;
                                    winrecoredStudent.paperType = winquestionWhich[d][m].paperType;
                                    winrecoredStudent.correctAnswer = winquestionWhich[d][m].correctAnswer;
                                    winrecoredStudentlist.Add(winrecoredStudent);
                                }
                            
                                winquestionbll.InsertWinRecordList(winrecoredStudentlist);
                            }
                        }                   
                   }
                }
            }

        //    //++++++++++++++++++++++++Excel+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            //3、ExcelQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable exceldt = new DataTable();
            exceldt = excelQuestionBll.ExcelPaperType();

            List<List<ExcelQuestionEntity>> excelquestionlist = new List<List<ExcelQuestionEntity>>();
            String excelPapertype = "";

            if (exceldt.Rows.Count != 0)
            {
                for (int i = 0; i < exceldt.Rows.Count; i++)
                {
                    excelPapertype = exceldt.Rows[i]["PaperType"].ToString();
                    excelquestionlist.Add(excelQuestionBll.ExcelPaperTypeGroupByPaperType(excelPapertype));
                }

                //4、给学生分配数据的PaperType,添加WordRecord表


                foreach (DataRow dr in dtcollege.Rows)
                {
                    collegeId = dr["collegeID"].ToString();
                    college.collegeID = collegeId;
                    if (examconfigbll.IsTableExist("ExcelQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("ExcelQuestionRecordEntity_" + collegeId, "ExcelQuestionRecordEntity");
                    }
                    List<StudentInfoEntity> studentinfolist = examconfigbll.GetStudentByCollege(college);

                    if (studentinfolist == null)
                    {
                        continue;
                    }

                    if (studentinfolist.Count > 0)
                    {

                        //进行分配
                        for (int j = 0; j < studentinfolist.Count; j++)
                        {
                            List<ExcelQuestionRecordEntity> excelrecoredStudentlist = new List<ExcelQuestionRecordEntity>();

                            String strStudentIDExcel = studentinfolist[j].studentID;
                            ExcelQuestionRecordEntity excelStudent = new ExcelQuestionRecordEntity();
                            excelStudent.StudentID = strStudentIDExcel;

                            if (excelQuestionBll.SelectExcelRecord(excelStudent) == true)
                            {
                                int i = excelquestionlist.Count;
                                int d = 0;
                                d = j % i;
                                for (int m = 0; m < excelquestionlist[d].Count; m++)
                                {
                                    ExcelQuestionRecordEntity excelrecoredStudent = new ExcelQuestionRecordEntity();
                                    excelrecoredStudent.StudentID = strStudentIDExcel;
                                    excelrecoredStudent.QuestionID = excelquestionlist[d][m].QuestionID;
                                    excelrecoredStudent.QuestionContent = excelquestionlist[d][m].QuestionContent;
                                    excelrecoredStudent.PaperType = excelquestionlist[d][m].PaperType;
                                    excelrecoredStudent.CorrectAnswer = excelquestionlist[d][m].CorrectAnswer;
                                    excelrecoredStudentlist.Add(excelrecoredStudent);
                                }
                                excelQuestionBll.InsertExcelRecordList(excelrecoredStudentlist);
                            }
                        }
                    }
                }

            }
        //    //++++++++++++++++++++++++IE+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            //IEQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable Ieddt = new DataTable();
            Ieddt = iequestionbll.WinPaperType();

            List<List<IEQuestionEntity>> iequestionWhich = new List<List<IEQuestionEntity>>();
            String iePapertype = "";

            if (Ieddt.Rows.Count != 0)
            {

                for (int i = 0; i < Ieddt.Rows.Count; i++)
                {
                    iePapertype = Ieddt.Rows[i]["paperType"].ToString();
                    iequestionWhich.Add(iequestionbll.IEPaperTypeGroupByPaperType(iePapertype));
                }




                foreach (DataRow dr in dtcollege.Rows)
                {
                    collegeId = dr["collegeID"].ToString();
                    college.collegeID = collegeId;
                    if (examconfigbll.IsTableExist("IEQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("IEQuestionRecordEntity_" + collegeId, "IEQuestionRecordEntity");
                    }
                    List<StudentInfoEntity> studentinfolist = examconfigbll.GetStudentByCollege(college);

                    if (studentinfolist == null)
                    {
                        continue;
                    }

                    if (studentinfolist.Count > 0)
                    {

                        //进行分配
                        for (int j = 0; j < studentinfolist.Count; j++)
                        {
                            //做循环
                            List<IEQuestionRecordEntity> ierecoredStudentlist = new List<IEQuestionRecordEntity>();

                            String strStudentIDIE = studentinfolist[j].studentID;
                            IEQuestionRecordEntity IEStudent = new IEQuestionRecordEntity();
                            IEStudent.studentID = strStudentIDIE;
                            if (iequestionbll.SelectIERecord(IEStudent) == true)
                            {

                                int i = iequestionWhich.Count;
                                int d = 0;
                                d = j % i;
                                for (int m = 0; m < iequestionWhich[d].Count; m++)
                                {
                                    IEQuestionRecordEntity ierecoredStudent = new IEQuestionRecordEntity();
                                    ierecoredStudent.studentID = strStudentIDIE;
                                    ierecoredStudent.questionID = iequestionWhich[d][m].questionID;
                                    ierecoredStudent.questionContent = iequestionWhich[d][m].questionContent;
                                    ierecoredStudent.paperType = iequestionWhich[d][m].paperType;
                                    ierecoredStudent.correctAnswer = iequestionWhich[d][m].correctAnswer;
                                    ierecoredStudentlist.Add(ierecoredStudent);
                                }
                                iequestionbll.InsertIERecordList(ierecoredStudentlist);
                            }
                        }
                    }
                }
            }

        //    //++++++++++++++++++++++++Ppt+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            //3、PptQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable pptdt = new DataTable();
            pptdt = pptQuestionBll.PptPaperType();

            List<List<PptQuestionEntity>> pptQuestionlist = new List<List<PptQuestionEntity>>();
            String PapertypePPT = "";
            if (pptdt.Rows.Count != 0)
            {
                for (int i = 0; i < pptdt.Rows.Count; i++)
                {
                    PapertypePPT = pptdt.Rows[i]["PaperType"].ToString();
                    pptQuestionlist.Add(pptQuestionBll.PptPaperTypeGroupByPaperType(PapertypePPT));
                }

                //4、给学生分配数据的PaperType,添加WordRecord表

                foreach (DataRow dr in dtcollege.Rows)
                {
                    collegeId = dr["collegeID"].ToString();
                    college.collegeID = collegeId;
                    if (examconfigbll.IsTableExist("PptQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("PptQuestionRecordEntity_" + collegeId, "PptQuestionRecordEntity");
                    }

                    List<StudentInfoEntity> studentinfolist = examconfigbll.GetStudentByCollege(college);

                    if (studentinfolist == null)
                    {
                        continue;
                    }

                    if (studentinfolist.Count > 0)
                    {
                        //进行分配
                        for (int j = 0; j < studentinfolist.Count; j++)
                        {
                            List<PptQuestionRecordEntity> pptrecoredStudentlist = new List<PptQuestionRecordEntity>();

                            String strStudentIDPpt = studentinfolist[j].studentID;
                            PptQuestionRecordEntity pptStudent = new PptQuestionRecordEntity();
                            pptStudent.StudentID = strStudentIDPpt;
                            if (pptQuestionBll.SelectPptRecord(pptStudent) == true)
                            {
                                int i = pptQuestionlist.Count;
                                int d = 0;
                                d = j % i;
                                for (int m = 0; m < pptQuestionlist[d].Count; m++)
                                {
                                    PptQuestionRecordEntity pptrecoredStudent = new PptQuestionRecordEntity();
                                    pptrecoredStudent.StudentID = strStudentIDPpt;
                                    pptrecoredStudent.QuestionID = pptQuestionlist[d][m].QuestionID;
                                    pptrecoredStudent.QuestionContent = pptQuestionlist[d][m].QuestionContent;
                                    pptrecoredStudent.PaperType = pptQuestionlist[d][m].PaperType;
                                    pptrecoredStudent.RightAnswer = pptQuestionlist[d][m].RightAnswer;
                                    pptrecoredStudentlist.Add(pptrecoredStudent);
                                }
                                pptQuestionBll.InsertPptRecordList(pptrecoredStudentlist);
                            }
                        }
                    }
                }
            }

           


            MessageBox.Show("配题成功！");
            //stopwatch.Stop(); //  停止监视
            //TimeSpan timespan = stopwatch.Elapsed; //  获取当前实例测量得出的总时间
            //double hours = timespan.TotalHours; // 总小时
            //double minutes = timespan.TotalMinutes;  // 总分钟
            //double seconds = timespan.TotalSeconds;  //  总秒数
            //double milliseconds = timespan.TotalMilliseconds;  //  总毫秒数
        }

 

        private void button3_Click(object sender, EventArgs e)
        {
            String collegeId;
            String studentID = SingleStudentConfigBox.Text.Trim();
            if (studentID != "")
            {
                StudentInfoEntity studentInfo = studentinfoBll.GetStudentById(studentID);

                //给该学生分配题
                 //++++++++++++++++++++++++Word+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                collegeId = studentInfo.CollegeID;
           

            //3、WordQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable worddt = new DataTable();
            worddt = wordquestionbll.WordPaperType();

            List<List<WordQuestionEntity>> wordquestionlist = new List<List<WordQuestionEntity>>();
            String Papertype = "";
            if (worddt.Rows.Count != 0)
            {
                for (int i = 0; i < worddt.Rows.Count; i++)
                {
                    Papertype = worddt.Rows[i]["PaperType"].ToString();
                    wordquestionlist.Add(wordquestionbll.WordPaperTypeGroupByPaperType(Papertype));
                }

                //4、给学生分配数据的PaperType,添加WordRecord表             
                   
                    if (examconfigbll.IsTableExist("WordQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("WordQuestionRecordEntity_" + collegeId, "WordQuestionRecordEntity");
                    }

                    List<StudentBindPaperTypeEntity> studentBindPaperTypeEntityList = new List<StudentBindPaperTypeEntity>();      //这里在StudentBindPaperTypeEntity中insert
 

             
                     List<WordQuestionRecordEntity> wordrecoredStudentlist = new List<WordQuestionRecordEntity>();

                        String strStudentID = studentInfo.studentID;
                        WordQuestionRecordEntity wordStudent = new WordQuestionRecordEntity();
                        wordStudent.StudentID = strStudentID;
                        if (wordquestionbll.SelectWordRecord(wordStudent) == true)
                        {
                            int i = wordquestionlist.Count;
                            int d = 0;
                            for (int m = 0; m < wordquestionlist[d].Count; m++)
                            {
                                WordQuestionRecordEntity wordrecoredStudent = new WordQuestionRecordEntity();
                                wordrecoredStudent.StudentID = strStudentID;                                  
                                wordrecoredStudent.QuestionID = wordquestionlist[d][m].QuestionID;
                                wordrecoredStudent.QuestionContent = wordquestionlist[d][m].QuestionContent;
                                wordrecoredStudent.PaperType = wordquestionlist[d][m].PaperType;
                                wordrecoredStudent.RightAnswer = wordquestionlist[d][m].RightAnswer;
                                wordrecoredStudentlist.Add(wordrecoredStudent);
                            }
                            wordquestionbll.InsertWordRecordList(wordrecoredStudentlist);


                            StudentBindPaperTypeEntity studentBindPaperTypeEntity = new StudentBindPaperTypeEntity();
                            studentBindPaperTypeEntity.StudentID = studentInfo.studentID;
                            studentBindPaperTypeEntity.CollegeID = studentInfo.CollegeID;
                            studentBindPaperTypeEntity.PaperType = wordquestionlist[d][0].PaperType;
                            studentBindPaperTypeEntity.IsUse = 1;
                            if (studentBindPaperType.SelectRecord(studentBindPaperTypeEntity) == true)
                            {
                                studentBindPaperTypeEntityList.Add(studentBindPaperTypeEntity);
                            }
                        }
                       
                        studentBindPaperType.InsertRecordList(studentBindPaperTypeEntityList);
                  }
   
   




            //++++++++++++++++++++++++Win+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


            //WinQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable winddt = new DataTable();
            winddt = winquestionbll.WinPaperType();

            List<List<WinQuestionEntity>> winquestionWhich = new List<List<WinQuestionEntity>>();

            String WinPapertype = "";

            if (winddt.Rows.Count != 0)
            { 
                                     
                for (int i = 0; i < winddt.Rows.Count; i++)
                {
                    WinPapertype = winddt.Rows[i]["paperType"].ToString();
                    winquestionWhich.Add(winquestionbll.WinPaperTypeGroupByPaperType(WinPapertype));
                }



                if (examconfigbll.IsTableExist("WinQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("WinQuestionRecordEntity_" + collegeId, "WinQuestionRecordEntity");
                    }

                  

                
                  
                        //进行分配
                        
                            //做循环
                            List<WinQuestionRecordEntity> winrecoredStudentlist = new List<WinQuestionRecordEntity>();

                            String StrStudentID = studentinfo.studentID;
                            WinQuestionRecordEntity winStudent = new WinQuestionRecordEntity();
                            winStudent.studentID = StrStudentID;
                            if (winquestionbll.SelectWinRecord(winStudent) == true)
                            {
                                int i = winquestionWhich.Count;
                                int d = 0;
                                for (int m = 0; m < winquestionWhich[d].Count; m++)
                                {
                                    WinQuestionRecordEntity winrecoredStudent = new WinQuestionRecordEntity();
                                    winrecoredStudent.studentID = StrStudentID;
                                    winrecoredStudent.questionID = winquestionWhich[d][m].questionID;
                                    winrecoredStudent.questionContent = winquestionWhich[d][m].questionContent;
                                    winrecoredStudent.paperType = winquestionWhich[d][m].paperType;
                                    winrecoredStudent.correctAnswer = winquestionWhich[d][m].correctAnswer;
                                    winrecoredStudentlist.Add(winrecoredStudent);
                                }
                            
                                winquestionbll.InsertWinRecordList(winrecoredStudentlist);
                            }
                                           
                
                }
          

            //++++++++++++++++++++++++Excel+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            //3、ExcelQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable exceldt = new DataTable();
            exceldt = excelQuestionBll.ExcelPaperType();

            List<List<ExcelQuestionEntity>> excelquestionlist = new List<List<ExcelQuestionEntity>>();
            String excelPapertype = "";

            if (exceldt.Rows.Count != 0)
            {              
                for (int i = 0; i < exceldt.Rows.Count; i++)
                {
                    excelPapertype = exceldt.Rows[i]["PaperType"].ToString();
                    excelquestionlist.Add(excelQuestionBll.ExcelPaperTypeGroupByPaperType(excelPapertype));
                }

                //4、给学生分配数据的PaperType,添加WordRecord表
                

             
                   
                    if (examconfigbll.IsTableExist("ExcelQuestionRecordEntity_" + collegeId)==false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("ExcelQuestionRecordEntity_" + collegeId,"ExcelQuestionRecordEntity");
                    }
                 

                

         
                   
                        //进行分配
                     
                            List<ExcelQuestionRecordEntity> excelrecoredStudentlist = new List<ExcelQuestionRecordEntity>();

                            String strStudentIDExcel = studentinfo.studentID;
                            ExcelQuestionRecordEntity excelStudent = new ExcelQuestionRecordEntity();
                            excelStudent.StudentID = strStudentIDExcel;

                            if (excelQuestionBll.SelectExcelRecord(excelStudent) == true) { 
                                int i = excelquestionlist.Count;
                                int d = 0;
                                for (int m = 0; m < excelquestionlist[d].Count; m++)
                                {
                                    ExcelQuestionRecordEntity excelrecoredStudent = new ExcelQuestionRecordEntity();
                                    excelrecoredStudent.StudentID = strStudentIDExcel;
                                    excelrecoredStudent.QuestionID = excelquestionlist[d][m].QuestionID;
                                    excelrecoredStudent.QuestionContent = excelquestionlist[d][m].QuestionContent;
                                    excelrecoredStudent.PaperType = excelquestionlist[d][m].PaperType;
                                    excelrecoredStudent.CorrectAnswer = excelquestionlist[d][m].CorrectAnswer;
                                    excelrecoredStudentlist.Add(excelrecoredStudent);
                                }
                                excelQuestionBll.InsertExcelRecordList(excelrecoredStudentlist);
                            }
                                     
  
              
                
            }
            //++++++++++++++++++++++++IE+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            //IEQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable Ieddt = new DataTable();
            Ieddt = iequestionbll.WinPaperType();

            List<List<IEQuestionEntity>> iequestionWhich = new List<List<IEQuestionEntity>>();
            String iePapertype = "";

            if (Ieddt.Rows.Count != 0)
            { 
                      
                for (int i = 0; i < Ieddt.Rows.Count; i++)
                {
                    iePapertype = Ieddt.Rows[i]["paperType"].ToString();
                    iequestionWhich.Add(iequestionbll.IEPaperTypeGroupByPaperType(iePapertype));
                }


            

         
                if (examconfigbll.IsTableExist("IEQuestionRecordEntity_" + collegeId)==false)
                {
                    examconfigbll.CreateDataTableCopySelectRecord("IEQuestionRecordEntity_" + collegeId, "IEQuestionRecordEntity");
                }

                

        
               
                    //进行分配
                  
                        //做循环
                        List<IEQuestionRecordEntity> ierecoredStudentlist = new List<IEQuestionRecordEntity>();

                        String strStudentIDIE = studentinfo.studentID;
                        IEQuestionRecordEntity IEStudent = new IEQuestionRecordEntity();
                        IEStudent.studentID = strStudentIDIE;
                        if (iequestionbll.SelectIERecord(IEStudent) == true)
                        {

                            int i = iequestionWhich.Count;
                            int d = 0;
                     
                            for (int m = 0; m < iequestionWhich[d].Count; m++)
                            {
                                IEQuestionRecordEntity ierecoredStudent = new IEQuestionRecordEntity();
                                ierecoredStudent.studentID = strStudentIDIE;
                                ierecoredStudent.questionID = iequestionWhich[d][m].questionID;
                                ierecoredStudent.questionContent = iequestionWhich[d][m].questionContent;
                                ierecoredStudent.paperType = iequestionWhich[d][m].paperType;
                                ierecoredStudent.correctAnswer = iequestionWhich[d][m].correctAnswer;
                                ierecoredStudentlist.Add(ierecoredStudent);
                            }                    
                            iequestionbll.InsertIERecordList(ierecoredStudentlist);
                        }
               
          
                
        }

            //++++++++++++++++++++++++Ppt+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            //3、PptQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
            DataTable pptdt = new DataTable();
            pptdt = pptQuestionBll.PptPaperType();
        
            List<List<PptQuestionEntity>> pptQuestionlist = new List<List<PptQuestionEntity>>();
            String PapertypePPT = "";
            if (pptdt.Rows.Count != 0)
            {
                for (int i = 0; i < pptdt.Rows.Count; i++)
                {
                    PapertypePPT = pptdt.Rows[i]["PaperType"].ToString();
                    pptQuestionlist.Add(pptQuestionBll.PptPaperTypeGroupByPaperType(PapertypePPT));
                }

                //4、给学生分配数据的PaperType,添加WordRecord表
                
              
                   
                    if (examconfigbll.IsTableExist("PptQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("PptQuestionRecordEntity_" + collegeId, "PptQuestionRecordEntity");
                    }

               


                                                                           
                        //进行分配
                       
                            List<PptQuestionRecordEntity> pptrecoredStudentlist = new List<PptQuestionRecordEntity>();
                           
                            String strStudentIDPpt = studentinfo.studentID;
                            PptQuestionRecordEntity pptStudent = new PptQuestionRecordEntity();
                            pptStudent.StudentID = strStudentIDPpt;
                            if (pptQuestionBll.SelectPptRecord(pptStudent) == true) 
                            {
                                int i = pptQuestionlist.Count;
                                int d = 0;
             
                                for (int m = 0; m < pptQuestionlist[d].Count; m++)
                                {
                                    PptQuestionRecordEntity pptrecoredStudent = new PptQuestionRecordEntity();
                                    pptrecoredStudent.StudentID = strStudentIDPpt;
                                    pptrecoredStudent.QuestionID = pptQuestionlist[d][m].QuestionID;
                                    pptrecoredStudent.QuestionContent = pptQuestionlist[d][m].QuestionContent;
                                    pptrecoredStudent.PaperType = pptQuestionlist[d][m].PaperType;
                                    pptrecoredStudent.RightAnswer = pptQuestionlist[d][m].RightAnswer;
                                    pptrecoredStudentlist.Add(pptrecoredStudent);
                                }                                   
                             pptQuestionBll.InsertPptRecordList(pptrecoredStudentlist);
                            }
                      
                    
              
                        }
                        MessageBox.Show("配题成功！");
            }
            else {
                MessageBox.Show("请填写学生的考号！");
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            String collegeId;
            String studentID = SingleStudentConfigBox.Text.Trim();
            if (studentID != "")
            {
                StudentInfoEntity studentInfo = studentinfoBll.GetStudentById(studentID);

                //给该学生分配题
                //++++++++++++++++++++++++Word+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                collegeId = studentInfo.CollegeID;


                //3、WordQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
                DataTable worddt = new DataTable();
                worddt = wordquestionbll.WordPaperType();

                List<List<WordQuestionEntity>> wordquestionlist = new List<List<WordQuestionEntity>>();
                String Papertype = "";
                if (worddt.Rows.Count != 0)
                {
                    for (int i = 0; i < worddt.Rows.Count; i++)
                    {
                        Papertype = worddt.Rows[i]["PaperType"].ToString();
                        wordquestionlist.Add(wordquestionbll.WordPaperTypeGroupByPaperType(Papertype));
                    }

                    //4、给学生分配数据的PaperType,添加WordRecord表             

                    if (examconfigbll.IsTableExist("WordQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("WordQuestionRecordEntity_" + collegeId, "WordQuestionRecordEntity");
                    }

                    List<StudentBindPaperTypeEntity> studentBindPaperTypeEntityList = new List<StudentBindPaperTypeEntity>();      //这里在StudentBindPaperTypeEntity中insert



                    List<WordQuestionRecordEntity> wordrecoredStudentlist = new List<WordQuestionRecordEntity>();

                    String strStudentID = studentInfo.studentID;
                    WordQuestionRecordEntity wordStudent = new WordQuestionRecordEntity();
                    wordStudent.StudentID = strStudentID;
                    if (wordquestionbll.SelectWordRecord(wordStudent) == true)
                    {
                        int i = wordquestionlist.Count;
                        int d = 0;
                        for (int m = 0; m < wordquestionlist[d].Count; m++)
                        {
                            WordQuestionRecordEntity wordrecoredStudent = new WordQuestionRecordEntity();
                            wordrecoredStudent.StudentID = strStudentID;
                            wordrecoredStudent.QuestionID = wordquestionlist[d][m].QuestionID;
                            wordrecoredStudent.QuestionContent = wordquestionlist[d][m].QuestionContent;
                            wordrecoredStudent.PaperType = wordquestionlist[d][m].PaperType;
                            wordrecoredStudent.RightAnswer = wordquestionlist[d][m].RightAnswer;
                            wordrecoredStudentlist.Add(wordrecoredStudent);
                        }
                        wordquestionbll.InsertWordRecordList(wordrecoredStudentlist);


                        StudentBindPaperTypeEntity studentBindPaperTypeEntity = new StudentBindPaperTypeEntity();
                        studentBindPaperTypeEntity.StudentID = studentInfo.studentID;
                        studentBindPaperTypeEntity.CollegeID = studentInfo.CollegeID;
                        studentBindPaperTypeEntity.PaperType = wordquestionlist[d][0].PaperType;
                        studentBindPaperTypeEntity.IsUse = 1;
                        if (studentBindPaperType.SelectRecord(studentBindPaperTypeEntity) == true)
                        {
                            studentBindPaperTypeEntityList.Add(studentBindPaperTypeEntity);
                        }
                    }
                    //-------------------------选择题配题-----------------------------------------------
                    //如果选中全部，则清空所有的选择题答题记录
                    //examConfigBll.ClearSelectQuestionRecordByCollegeID(allCollege);
                    examconfigbll.ClearSelectQuestionRecordByLstStudent(new List<StudentInfoEntity>() { studentInfo});
                    examconfigbll.RandGenerateRecord(new List<StudentInfoEntity>(){studentInfo});

                    studentBindPaperType.InsertRecordList(studentBindPaperTypeEntityList);
                }






                //++++++++++++++++++++++++Win+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


                //WinQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
                DataTable winddt = new DataTable();
                winddt = winquestionbll.WinPaperType();

                List<List<WinQuestionEntity>> winquestionWhich = new List<List<WinQuestionEntity>>();

                String WinPapertype = "";

                if (winddt.Rows.Count != 0)
                {

                    for (int i = 0; i < winddt.Rows.Count; i++)
                    {
                        WinPapertype = winddt.Rows[i]["paperType"].ToString();
                        winquestionWhich.Add(winquestionbll.WinPaperTypeGroupByPaperType(WinPapertype));
                    }



                    if (examconfigbll.IsTableExist("WinQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("WinQuestionRecordEntity_" + collegeId, "WinQuestionRecordEntity");
                    }





                    //进行分配

                    //做循环
                    List<WinQuestionRecordEntity> winrecoredStudentlist = new List<WinQuestionRecordEntity>();

                    String StrStudentID = studentInfo.studentID;
                    WinQuestionRecordEntity winStudent = new WinQuestionRecordEntity();
                    winStudent.studentID = StrStudentID;
                    if (winquestionbll.SelectWinRecord(winStudent) == true)
                    {
                        int i = winquestionWhich.Count;
                        int d = 0;
                        for (int m = 0; m < winquestionWhich[d].Count; m++)
                        {
                            WinQuestionRecordEntity winrecoredStudent = new WinQuestionRecordEntity();
                            winrecoredStudent.studentID = StrStudentID;
                            winrecoredStudent.questionID = winquestionWhich[d][m].questionID;
                            winrecoredStudent.questionContent = winquestionWhich[d][m].questionContent;
                            winrecoredStudent.paperType = winquestionWhich[d][m].paperType;
                            winrecoredStudent.correctAnswer = winquestionWhich[d][m].correctAnswer;
                            winrecoredStudentlist.Add(winrecoredStudent);
                        }

                        winquestionbll.InsertWinRecordList(winrecoredStudentlist);
                    }


                }


                //++++++++++++++++++++++++Excel+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                //3、ExcelQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
                DataTable exceldt = new DataTable();
                exceldt = excelQuestionBll.ExcelPaperType();

                List<List<ExcelQuestionEntity>> excelquestionlist = new List<List<ExcelQuestionEntity>>();
                String excelPapertype = "";

                if (exceldt.Rows.Count != 0)
                {
                    for (int i = 0; i < exceldt.Rows.Count; i++)
                    {
                        excelPapertype = exceldt.Rows[i]["PaperType"].ToString();
                        excelquestionlist.Add(excelQuestionBll.ExcelPaperTypeGroupByPaperType(excelPapertype));
                    }

                    //4、给学生分配数据的PaperType,添加WordRecord表




                    if (examconfigbll.IsTableExist("ExcelQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("ExcelQuestionRecordEntity_" + collegeId, "ExcelQuestionRecordEntity");
                    }






                    //进行分配

                    List<ExcelQuestionRecordEntity> excelrecoredStudentlist = new List<ExcelQuestionRecordEntity>();

                    String strStudentIDExcel = studentInfo.studentID;
                    ExcelQuestionRecordEntity excelStudent = new ExcelQuestionRecordEntity();
                    excelStudent.StudentID = strStudentIDExcel;

                    if (excelQuestionBll.SelectExcelRecord(excelStudent) == true)
                    {
                        int i = excelquestionlist.Count;
                        int d = 0;
                        for (int m = 0; m < excelquestionlist[d].Count; m++)
                        {
                            ExcelQuestionRecordEntity excelrecoredStudent = new ExcelQuestionRecordEntity();
                            excelrecoredStudent.StudentID = strStudentIDExcel;
                            excelrecoredStudent.QuestionID = excelquestionlist[d][m].QuestionID;
                            excelrecoredStudent.QuestionContent = excelquestionlist[d][m].QuestionContent;
                            excelrecoredStudent.PaperType = excelquestionlist[d][m].PaperType;
                            excelrecoredStudent.CorrectAnswer = excelquestionlist[d][m].CorrectAnswer;
                            excelrecoredStudentlist.Add(excelrecoredStudent);
                        }
                        excelQuestionBll.InsertExcelRecordList(excelrecoredStudentlist);
                    }




                }
                //++++++++++++++++++++++++IE+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                //IEQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
                DataTable Ieddt = new DataTable();
                Ieddt = iequestionbll.WinPaperType();

                List<List<IEQuestionEntity>> iequestionWhich = new List<List<IEQuestionEntity>>();
                String iePapertype = "";

                if (Ieddt.Rows.Count != 0)
                {

                    for (int i = 0; i < Ieddt.Rows.Count; i++)
                    {
                        iePapertype = Ieddt.Rows[i]["paperType"].ToString();
                        iequestionWhich.Add(iequestionbll.IEPaperTypeGroupByPaperType(iePapertype));
                    }





                    if (examconfigbll.IsTableExist("IEQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("IEQuestionRecordEntity_" + collegeId, "IEQuestionRecordEntity");
                    }





                    //进行分配

                    //做循环
                    List<IEQuestionRecordEntity> ierecoredStudentlist = new List<IEQuestionRecordEntity>();

                    String strStudentIDIE = studentInfo.studentID;
                    IEQuestionRecordEntity IEStudent = new IEQuestionRecordEntity();
                    IEStudent.studentID = strStudentIDIE;
                    if (iequestionbll.SelectIERecord(IEStudent) == true)
                    {

                        int i = iequestionWhich.Count;
                        int d = 0;

                        for (int m = 0; m < iequestionWhich[d].Count; m++)
                        {
                            IEQuestionRecordEntity ierecoredStudent = new IEQuestionRecordEntity();
                            ierecoredStudent.studentID = strStudentIDIE;
                            ierecoredStudent.questionID = iequestionWhich[d][m].questionID;
                            ierecoredStudent.questionContent = iequestionWhich[d][m].questionContent;
                            ierecoredStudent.paperType = iequestionWhich[d][m].paperType;
                            ierecoredStudent.correctAnswer = iequestionWhich[d][m].correctAnswer;
                            ierecoredStudentlist.Add(ierecoredStudent);
                        }
                        iequestionbll.InsertIERecordList(ierecoredStudentlist);
                    }



                }

                //++++++++++++++++++++++++Ppt+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                //3、PptQuestionEntity中查找出所有的信息。将不同的ABCd等试卷的信息存放在一个list中
                DataTable pptdt = new DataTable();
                pptdt = pptQuestionBll.PptPaperType();

                List<List<PptQuestionEntity>> pptQuestionlist = new List<List<PptQuestionEntity>>();
                String PapertypePPT = "";
                if (pptdt.Rows.Count != 0)
                {
                    for (int i = 0; i < pptdt.Rows.Count; i++)
                    {
                        PapertypePPT = pptdt.Rows[i]["PaperType"].ToString();
                        pptQuestionlist.Add(pptQuestionBll.PptPaperTypeGroupByPaperType(PapertypePPT));
                    }

                    //4、给学生分配数据的PaperType,添加WordRecord表



                    if (examconfigbll.IsTableExist("PptQuestionRecordEntity_" + collegeId) == false)
                    {
                        examconfigbll.CreateDataTableCopySelectRecord("PptQuestionRecordEntity_" + collegeId, "PptQuestionRecordEntity");
                    }





                    //进行分配

                    List<PptQuestionRecordEntity> pptrecoredStudentlist = new List<PptQuestionRecordEntity>();

                    String strStudentIDPpt = studentInfo.studentID;
                    PptQuestionRecordEntity pptStudent = new PptQuestionRecordEntity();
                    pptStudent.StudentID = strStudentIDPpt;
                    if (pptQuestionBll.SelectPptRecord(pptStudent) == true)
                    {
                        int i = pptQuestionlist.Count;
                        int d = 0;

                        for (int m = 0; m < pptQuestionlist[d].Count; m++)
                        {
                            PptQuestionRecordEntity pptrecoredStudent = new PptQuestionRecordEntity();
                            pptrecoredStudent.StudentID = strStudentIDPpt;
                            pptrecoredStudent.QuestionID = pptQuestionlist[d][m].QuestionID;
                            pptrecoredStudent.QuestionContent = pptQuestionlist[d][m].QuestionContent;
                            pptrecoredStudent.PaperType = pptQuestionlist[d][m].PaperType;
                            pptrecoredStudent.RightAnswer = pptQuestionlist[d][m].RightAnswer;
                            pptrecoredStudentlist.Add(pptrecoredStudent);
                        }
                        pptQuestionBll.InsertPptRecordList(pptrecoredStudentlist);
                    }



                }
                MessageBox.Show("配题成功！");
            }
            else
            {
                MessageBox.Show("请填写学生的考号！");
            }
        }
    }
}
