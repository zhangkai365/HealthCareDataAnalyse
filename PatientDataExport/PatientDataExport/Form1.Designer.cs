namespace PatientDataExport
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_beginProgress = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_selectSavePath = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtbox_FilePath = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.processOutputExcel = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.processStatistics = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.iffinished = new System.Windows.Forms.Label();
            this.totalNum = new System.Windows.Forms.Label();
            this.progressNum = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.datePicker_startDate = new System.Windows.Forms.DateTimePicker();
            this.lab_endDate = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.chk_CreateNewDiseaseList = new System.Windows.Forms.CheckBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.txt_WorkUnit = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_beginProgress
            // 
            this.btn_beginProgress.Location = new System.Drawing.Point(527, 464);
            this.btn_beginProgress.Name = "btn_beginProgress";
            this.btn_beginProgress.Size = new System.Drawing.Size(144, 32);
            this.btn_beginProgress.TabIndex = 2;
            this.btn_beginProgress.Text = "开始导出数据";
            this.btn_beginProgress.UseVisualStyleBackColor = true;
            this.btn_beginProgress.Click += new System.EventHandler(this.btn_beginProgress_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_selectSavePath);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtbox_FilePath);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(890, 76);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "文件保存路径";
            // 
            // btn_selectSavePath
            // 
            this.btn_selectSavePath.Location = new System.Drawing.Point(802, 20);
            this.btn_selectSavePath.Name = "btn_selectSavePath";
            this.btn_selectSavePath.Size = new System.Drawing.Size(71, 44);
            this.btn_selectSavePath.TabIndex = 12;
            this.btn_selectSavePath.Text = "选择Excel文件存储路径";
            this.btn_selectSavePath.UseVisualStyleBackColor = true;
            this.btn_selectSavePath.Click += new System.EventHandler(this.btn_selectSavePath_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(-131, 34);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 12);
            this.label3.TabIndex = 11;
            this.label3.Text = "Excel文件存储路径";
            // 
            // txtbox_FilePath
            // 
            this.txtbox_FilePath.Location = new System.Drawing.Point(15, 33);
            this.txtbox_FilePath.Name = "txtbox_FilePath";
            this.txtbox_FilePath.Size = new System.Drawing.Size(772, 21);
            this.txtbox_FilePath.TabIndex = 10;
            this.txtbox_FilePath.Text = "C:\\Users\\win7x64_20150617\\Desktop\\20150721PatientAnalyse\\市委宣传部07210006.xlsx";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.processOutputExcel);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.processStatistics);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.iffinished);
            this.groupBox2.Controls.Add(this.totalNum);
            this.groupBox2.Controls.Add(this.progressNum);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(457, 285);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(445, 134);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "文件处理进度";
            // 
            // processOutputExcel
            // 
            this.processOutputExcel.AutoSize = true;
            this.processOutputExcel.Location = new System.Drawing.Point(156, 91);
            this.processOutputExcel.Name = "processOutputExcel";
            this.processOutputExcel.Size = new System.Drawing.Size(41, 12);
            this.processOutputExcel.TabIndex = 15;
            this.processOutputExcel.Text = "未开始";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(22, 91);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(119, 12);
            this.label8.TabIndex = 14;
            this.label8.Text = "输出Excel执行过程：";
            // 
            // processStatistics
            // 
            this.processStatistics.AutoSize = true;
            this.processStatistics.Location = new System.Drawing.Point(156, 62);
            this.processStatistics.Name = "processStatistics";
            this.processStatistics.Size = new System.Drawing.Size(41, 12);
            this.processStatistics.TabIndex = 13;
            this.processStatistics.Text = "未开始";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 62);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(113, 12);
            this.label6.TabIndex = 12;
            this.label6.Text = "统计过程执行情况：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(144, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(11, 12);
            this.label1.TabIndex = 11;
            this.label1.Text = "/";
            // 
            // iffinished
            // 
            this.iffinished.AutoSize = true;
            this.iffinished.Location = new System.Drawing.Point(280, 25);
            this.iffinished.Name = "iffinished";
            this.iffinished.Size = new System.Drawing.Size(41, 12);
            this.iffinished.TabIndex = 10;
            this.iffinished.Text = "未完成";
            // 
            // totalNum
            // 
            this.totalNum.AutoSize = true;
            this.totalNum.Location = new System.Drawing.Point(176, 25);
            this.totalNum.Name = "totalNum";
            this.totalNum.Size = new System.Drawing.Size(11, 12);
            this.totalNum.TabIndex = 9;
            this.totalNum.Text = "0";
            // 
            // progressNum
            // 
            this.progressNum.AutoSize = true;
            this.progressNum.Location = new System.Drawing.Point(106, 25);
            this.progressNum.Name = "progressNum";
            this.progressNum.Size = new System.Drawing.Size(11, 12);
            this.progressNum.TabIndex = 8;
            this.progressNum.Text = "0";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "进度：";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.datePicker_startDate);
            this.groupBox3.Controls.Add(this.lab_endDate);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Location = new System.Drawing.Point(579, 20);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(294, 134);
            this.groupBox3.TabIndex = 13;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "处理数据时间范围";
            // 
            // datePicker_startDate
            // 
            this.datePicker_startDate.Location = new System.Drawing.Point(126, 35);
            this.datePicker_startDate.Name = "datePicker_startDate";
            this.datePicker_startDate.Size = new System.Drawing.Size(147, 21);
            this.datePicker_startDate.TabIndex = 13;
            this.datePicker_startDate.Value = new System.DateTime(2015, 1, 1, 0, 0, 0, 0);
            this.datePicker_startDate.ValueChanged += new System.EventHandler(this.datePicker_startDate_ValueChanged);
            // 
            // lab_endDate
            // 
            this.lab_endDate.AutoSize = true;
            this.lab_endDate.Location = new System.Drawing.Point(124, 90);
            this.lab_endDate.Name = "lab_endDate";
            this.lab_endDate.Size = new System.Drawing.Size(89, 12);
            this.lab_endDate.TabIndex = 3;
            this.lab_endDate.Text = "2014年12月31日";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(28, 90);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 1;
            this.label5.Text = "截止日期：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(28, 41);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 0;
            this.label4.Text = "起始日期：";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.chk_CreateNewDiseaseList);
            this.groupBox4.Location = new System.Drawing.Point(13, 285);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(421, 66);
            this.groupBox4.TabIndex = 14;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "是否要建立新的疾病列表";
            // 
            // chk_CreateNewDiseaseList
            // 
            this.chk_CreateNewDiseaseList.AutoSize = true;
            this.chk_CreateNewDiseaseList.Checked = true;
            this.chk_CreateNewDiseaseList.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk_CreateNewDiseaseList.Location = new System.Drawing.Point(41, 30);
            this.chk_CreateNewDiseaseList.Name = "chk_CreateNewDiseaseList";
            this.chk_CreateNewDiseaseList.Size = new System.Drawing.Size(120, 16);
            this.chk_CreateNewDiseaseList.TabIndex = 0;
            this.chk_CreateNewDiseaseList.Text = "建立新的疾病列表";
            this.chk_CreateNewDiseaseList.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.txt_WorkUnit);
            this.groupBox5.Controls.Add(this.label7);
            this.groupBox5.Controls.Add(this.groupBox3);
            this.groupBox5.Location = new System.Drawing.Point(13, 103);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(889, 162);
            this.groupBox5.TabIndex = 15;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "统计范围";
            // 
            // txt_WorkUnit
            // 
            this.txt_WorkUnit.Location = new System.Drawing.Point(96, 20);
            this.txt_WorkUnit.Name = "txt_WorkUnit";
            this.txt_WorkUnit.Size = new System.Drawing.Size(177, 21);
            this.txt_WorkUnit.TabIndex = 1;
            this.txt_WorkUnit.Text = "天津市委宣传部";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(25, 24);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 0;
            this.label7.Text = "查体单位：";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(914, 608);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btn_beginProgress);
            this.Name = "Form1";
            this.Text = "体检数据导出为Excel格式";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_beginProgress;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn_selectSavePath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtbox_FilePath;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label iffinished;
        private System.Windows.Forms.Label totalNum;
        private System.Windows.Forms.Label progressNum;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label lab_endDate;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker datePicker_startDate;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckBox chk_CreateNewDiseaseList;
        private System.Windows.Forms.Label processOutputExcel;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label processStatistics;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txt_WorkUnit;
    }
}

