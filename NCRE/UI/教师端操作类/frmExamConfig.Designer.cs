namespace NCRE学生考试端V1._0
{
    partial class frmExamConfig
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmExamConfig));
            this.btnClear = new System.Windows.Forms.Button();
            this.cbCollege = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnConfigRecord = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dgvStudent = new System.Windows.Forms.DataGridView();
            this.isCheck = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.studentID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.studentName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.major = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.grade = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sex = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.majorClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvStudent)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(30, 400);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 0;
            this.btnClear.Text = "清空题库";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // cbCollege
            // 
            this.cbCollege.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbCollege.FormattingEnabled = true;
            this.cbCollege.Location = new System.Drawing.Point(29, 54);
            this.cbCollege.Name = "cbCollege";
            this.cbCollege.Size = new System.Drawing.Size(121, 20);
            this.cbCollege.TabIndex = 1;
            this.cbCollege.SelectedIndexChanged += new System.EventHandler(this.cbCollege_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "学院";
            // 
            // btnConfigRecord
            // 
            this.btnConfigRecord.Location = new System.Drawing.Point(643, 400);
            this.btnConfigRecord.Name = "btnConfigRecord";
            this.btnConfigRecord.Size = new System.Drawing.Size(75, 23);
            this.btnConfigRecord.TabIndex = 3;
            this.btnConfigRecord.Text = "配置考试";
            this.btnConfigRecord.UseVisualStyleBackColor = true;
            this.btnConfigRecord.Click += new System.EventHandler(this.btnConfigRecord_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 429);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(311, 36);
            this.label2.TabIndex = 4;
            this.label2.Text = "清空题库：根据所选学院进行答题记录的清空。\r\n      1、清空选定学生的答题记录信息\r\n      2、如果未选中学生，默认清空整个学院的答题记录";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(431, 429);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(287, 36);
            this.label3.TabIndex = 5;
            this.label3.Text = "配置考试：随机抽题（选择题）生成答题记录表\r\n      1、对选中学生进行抽题。\r\n      2、如果未选中学生，默认对整个学院进行抽题";
            // 
            // dgvStudent
            // 
            this.dgvStudent.AllowUserToAddRows = false;
            this.dgvStudent.AllowUserToDeleteRows = false;
            this.dgvStudent.AllowUserToResizeColumns = false;
            this.dgvStudent.AllowUserToResizeRows = false;
            this.dgvStudent.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dgvStudent.ColumnHeadersHeight = 25;
            this.dgvStudent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvStudent.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.isCheck,
            this.studentID,
            this.studentName,
            this.major,
            this.grade,
            this.sex,
            this.majorClass});
            this.dgvStudent.Location = new System.Drawing.Point(29, 92);
            this.dgvStudent.Name = "dgvStudent";
            this.dgvStudent.ReadOnly = true;
            this.dgvStudent.RowHeadersVisible = false;
            this.dgvStudent.RowHeadersWidth = 4;
            this.dgvStudent.RowTemplate.Height = 23;
            this.dgvStudent.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvStudent.Size = new System.Drawing.Size(689, 294);
            this.dgvStudent.TabIndex = 6;
            this.dgvStudent.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvStudent_CellClick);
            // 
            // isCheck
            // 
            this.isCheck.HeaderText = " ";
            this.isCheck.Name = "isCheck";
            this.isCheck.ReadOnly = true;
            this.isCheck.Width = 30;
            // 
            // studentID
            // 
            this.studentID.DataPropertyName = "studentID";
            this.studentID.HeaderText = "学号";
            this.studentID.MinimumWidth = 4;
            this.studentID.Name = "studentID";
            this.studentID.ReadOnly = true;
            this.studentID.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // studentName
            // 
            this.studentName.DataPropertyName = "studentName";
            this.studentName.HeaderText = "姓名";
            this.studentName.Name = "studentName";
            this.studentName.ReadOnly = true;
            this.studentName.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // major
            // 
            this.major.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.major.DataPropertyName = "major";
            this.major.HeaderText = "专业";
            this.major.Name = "major";
            this.major.ReadOnly = true;
            this.major.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // grade
            // 
            this.grade.DataPropertyName = "grade";
            this.grade.HeaderText = "年级";
            this.grade.Name = "grade";
            this.grade.ReadOnly = true;
            this.grade.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.grade.Width = 80;
            // 
            // sex
            // 
            this.sex.DataPropertyName = "sex";
            this.sex.HeaderText = "性别";
            this.sex.Name = "sex";
            this.sex.ReadOnly = true;
            this.sex.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.sex.Width = 50;
            // 
            // majorClass
            // 
            this.majorClass.DataPropertyName = "majorClass";
            this.majorClass.HeaderText = "主修课";
            this.majorClass.Name = "majorClass";
            this.majorClass.ReadOnly = true;
            this.majorClass.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.majorClass.Width = 120;
            // 
            // frmExamConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(753, 481);
            this.Controls.Add(this.dgvStudent);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnConfigRecord);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbCollege);
            this.Controls.Add(this.btnClear);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmExamConfig";
            this.Text = "考试配置";
            this.Load += new System.EventHandler(this.frmExamConfig_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvStudent)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.ComboBox cbCollege;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnConfigRecord;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView dgvStudent;
        private System.Windows.Forms.DataGridViewCheckBoxColumn isCheck;
        private System.Windows.Forms.DataGridViewTextBoxColumn studentID;
        private System.Windows.Forms.DataGridViewTextBoxColumn studentName;
        private System.Windows.Forms.DataGridViewTextBoxColumn major;
        private System.Windows.Forms.DataGridViewTextBoxColumn grade;
        private System.Windows.Forms.DataGridViewTextBoxColumn sex;
        private System.Windows.Forms.DataGridViewTextBoxColumn majorClass;
    }
}