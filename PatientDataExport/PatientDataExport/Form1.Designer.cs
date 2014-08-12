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
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_beginProgress
            // 
            this.btn_beginProgress.Location = new System.Drawing.Point(525, 274);
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
            this.groupBox1.Size = new System.Drawing.Size(802, 134);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "文件保存路径";
            // 
            // btn_selectSavePath
            // 
            this.btn_selectSavePath.Location = new System.Drawing.Point(298, 74);
            this.btn_selectSavePath.Name = "btn_selectSavePath";
            this.btn_selectSavePath.Size = new System.Drawing.Size(144, 31);
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
            this.txtbox_FilePath.Location = new System.Drawing.Point(11, 35);
            this.txtbox_FilePath.Name = "txtbox_FilePath";
            this.txtbox_FilePath.Size = new System.Drawing.Size(772, 21);
            this.txtbox_FilePath.TabIndex = 10;
            this.txtbox_FilePath.Text = "C:\\Users\\Win7x64_20140606\\Desktop\\2014年查体工作总结（修改201408080753）\\2014统计结论\\正局级0357.xl" +
    "s";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.iffinished);
            this.groupBox2.Controls.Add(this.totalNum);
            this.groupBox2.Controls.Add(this.progressNum);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(369, 183);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(445, 57);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "文件处理进度";
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
            this.groupBox3.Location = new System.Drawing.Point(12, 172);
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
            this.datePicker_startDate.Value = new System.DateTime(2014, 1, 1, 0, 0, 0, 0);
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
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(842, 322);
            this.Controls.Add(this.groupBox3);
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
    }
}

