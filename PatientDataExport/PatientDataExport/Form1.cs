using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

//include
using PatientDataExport;
using PatientDataExport.Data;
using Excel = Microsoft.Office.Interop.Excel;


namespace PatientDataExport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_beginProgress_Click(object sender, EventArgs e)
        {
            //禁用界面上面所有按钮
            btn_beginProgress.Enabled = false;
            btn_selectSavePath.Enabled = false;
            datePicker_startDate.Enabled = false;

            //设置要查询的时间
            DateTime startDate;
            startDate = datePicker_startDate.Value;
            //控制输入的日期的有效值
            DateTime endDate;
            endDate = datePicker_startDate.Value.AddYears(1);
            lab_endDate.Text = endDate.ToString();
            //文件的存储路径
            String FilePath = txtbox_FilePath.Text;
            Excel.Application myExcel = new Excel.Application();
            myExcel.Visible = false;
            //存储统计结果
            Excel.Workbook myWorkbook = myExcel.Workbooks.Add(true);
            Excel.Worksheet myWorkSheet = myWorkbook.Worksheets[1];

            //存储主要的疾病的ICD诊断号码的Excel
            Excel.Workbook ICDWorkbook = myExcel.Workbooks.Add(true);
            Excel.Worksheet ICDWorksheet = ICDWorkbook.Worksheets[1];

            //全部疾病诊断列表名称
            Dictionary<string, string> Dic_DiseaseList = new Dictionary<string, string>();
            DiseaseList myDiseaseList = new DiseaseList();
            //读取列表
            myDiseaseList.Initialize("", out Dic_DiseaseList);

            //DiseaseList myDiseaseList = new DiseaseList();
            //myDiseaseList.Initialize("", out List_Disease);

            //StringBuilder sb = new StringBuilder();
            //foreach (var temp in List_Disease)
            //{
            //    sb.Append(temp.Key);
            //    sb.Append("%");
            //    sb.Append(temp.Value);
            //    sb.Append("/n");
            //}
            //MessageBox.Show(sb.ToString());
            //所有疾病的查询字典
            Dic myDic = new Dic();

            //进行统计
            processStatistics.Text = "统计开始";
            StatisticsService myStatistics = new StatisticsService();
            processStatistics.Text =  myStatistics.statistics(startDate, endDate, ref myDic);
            //进行输出结果
            processOutputExcel.Text = "输出Excel开始";
            OutputExcelService myOutputExcel = new OutputExcelService();
            processOutputExcel.Text = myOutputExcel.OutputExcel(FilePath, myDic,Dic_DiseaseList);

        }

        private void btn_selectSavePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog myFileDialog = new OpenFileDialog();
            myFileDialog.Filter = "Excel|*.xls";
            if (myFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtbox_FilePath.Text = myFileDialog.FileName;
            }
        }

        //界面的年份选取值的限制
        private void datePicker_startDate_ValueChanged(object sender, EventArgs e)
        {
            if (datePicker_startDate.Value > Convert.ToDateTime("2015-1-1 00:00:00")) datePicker_startDate.Value = Convert.ToDateTime("2015-1-1 00:00:00");
            if (datePicker_startDate.Value < Convert.ToDateTime("2008-1-1 00:00:00")) datePicker_startDate.Value = Convert.ToDateTime("2008-1-1 00:00:00");
            lab_endDate.Text = datePicker_startDate.Value.AddYears(1).ToShortDateString();
        }

    }
}
