namespace NCRE学生考试端V1._0
{
    partial class frmTeacherManagerMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTeacherManagerMain));
            this.btnCalssScore = new System.Windows.Forms.Button();
            this.dgvScore = new System.Windows.Forms.DataGridView();
            this.cboCollege = new System.Windows.Forms.ComboBox();
            this.collegeEntityBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.cboMajor = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button3 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.SingleStudentConfigBox = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.btnCollegeScore = new System.Windows.Forms.Button();
            this.directorySearcher1 = new System.DirectoryServices.DirectorySearcher();
            ((System.ComponentModel.ISupportInitialize)(this.dgvScore)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.collegeEntityBindingSource)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCalssScore
            // 
            this.btnCalssScore.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnCalssScore.Location = new System.Drawing.Point(25, 24);
            this.btnCalssScore.Name = "btnCalssScore";
            this.btnCalssScore.Size = new System.Drawing.Size(264, 87);
            this.btnCalssScore.TabIndex = 0;
            this.btnCalssScore.Text = "一键导出所有学生成绩    （以班级为单位）";
            this.btnCalssScore.UseVisualStyleBackColor = true;
            this.btnCalssScore.Click += new System.EventHandler(this.btnCalssScore_Click);
            // 
            // dgvScore
            // 
            this.dgvScore.AllowUserToAddRows = false;
            this.dgvScore.AllowUserToDeleteRows = false;
            this.dgvScore.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvScore.Location = new System.Drawing.Point(344, 106);
            this.dgvScore.Name = "dgvScore";
            this.dgvScore.RowTemplate.Height = 23;
            this.dgvScore.Size = new System.Drawing.Size(856, 525);
            this.dgvScore.TabIndex = 1;
            this.dgvScore.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvScore_CellContentClick);
            // 
            // cboCollege
            // 
            this.cboCollege.DataSource = this.collegeEntityBindingSource;
            this.cboCollege.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboCollege.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboCollege.FormattingEnabled = true;
            this.cboCollege.Location = new System.Drawing.Point(435, 22);
            this.cboCollege.Name = "cboCollege";
            this.cboCollege.Size = new System.Drawing.Size(232, 24);
            this.cboCollege.TabIndex = 2;
            this.cboCollege.SelectedIndexChanged += new System.EventHandler(this.cboCollege_SelectedIndexChanged);
            // 
            // cboMajor
            // 
            this.cboMajor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMajor.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboMajor.FormattingEnabled = true;
            this.cboMajor.Location = new System.Drawing.Point(788, 22);
            this.cboMajor.Name = "cboMajor";
            this.cboMajor.Size = new System.Drawing.Size(230, 24);
            this.cboMajor.TabIndex = 3;
            this.cboMajor.SelectedIndexChanged += new System.EventHandler(this.cboMajor_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(341, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 16);
            this.label1.TabIndex = 4;
            this.label1.Text = "选择学院：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(694, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 16);
            this.label2.TabIndex = 5;
            this.label2.Text = "选择专业：";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.SingleStudentConfigBox);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.btnCollegeScore);
            this.panel1.Controls.Add(this.btnCalssScore);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(316, 659);
            this.panel1.TabIndex = 6;
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button3.Location = new System.Drawing.Point(25, 571);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(264, 59);
            this.button3.TabIndex = 6;
            this.button3.Text = "单个学生配题";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(11, 519);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(133, 19);
            this.label3.TabIndex = 5;
            this.label3.Text = "输入学生考号:";
            // 
            // SingleStudentConfigBox
            // 
            this.SingleStudentConfigBox.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SingleStudentConfigBox.Location = new System.Drawing.Point(150, 514);
            this.SingleStudentConfigBox.Name = "SingleStudentConfigBox";
            this.SingleStudentConfigBox.Size = new System.Drawing.Size(151, 31);
            this.SingleStudentConfigBox.TabIndex = 4;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button2.Location = new System.Drawing.Point(25, 330);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(264, 83);
            this.button2.TabIndex = 3;
            this.button2.Text = "一键配题";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnCollegeScore
            // 
            this.btnCollegeScore.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnCollegeScore.Location = new System.Drawing.Point(25, 152);
            this.btnCollegeScore.Name = "btnCollegeScore";
            this.btnCollegeScore.Size = new System.Drawing.Size(264, 87);
            this.btnCollegeScore.TabIndex = 1;
            this.btnCollegeScore.Text = "一键导出所有学生成绩     （以学院为单位）";
            this.btnCollegeScore.UseVisualStyleBackColor = true;
            this.btnCollegeScore.Click += new System.EventHandler(this.btnCollegeScore_Click);
            // 
            // directorySearcher1
            // 
            this.directorySearcher1.ClientTimeout = System.TimeSpan.Parse("-00:00:01");
            this.directorySearcher1.ServerPageTimeLimit = System.TimeSpan.Parse("-00:00:01");
            this.directorySearcher1.ServerTimeLimit = System.TimeSpan.Parse("-00:00:01");
            // 
            // frmTeacherManagerMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1223, 659);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboMajor);
            this.Controls.Add(this.cboCollege);
            this.Controls.Add(this.dgvScore);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmTeacherManagerMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "导出学生成绩";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.frmTeacherManagerMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvScore)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.collegeEntityBindingSource)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCalssScore;
        private System.Windows.Forms.DataGridView dgvScore;
        private System.Windows.Forms.ComboBox cboCollege;
        private System.Windows.Forms.BindingSource collegeEntityBindingSource;
        private System.Windows.Forms.ComboBox cboMajor;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
        private System.DirectoryServices.DirectorySearcher directorySearcher1;
        private System.Windows.Forms.Button btnCollegeScore;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox SingleStudentConfigBox;
    }
}