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
using PatientDataExport.Package;


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
            if (txt_WorkUnit.Text == "") return;
            //要统计的条件
            StatisticsParameters mnStatisticsParameters = new StatisticsParameters();
            mnStatisticsParameters.startDate = datePicker_startDate.Value;
            lab_endDate.Text = datePicker_startDate.Value.AddYears(1).ToString();
            mnStatisticsParameters.endDate = datePicker_startDate.Value.AddYears(1);
            mnStatisticsParameters.workunit = txt_WorkUnit.Text.ToString();
            
            //禁用界面上面所有按钮
            btn_beginProgress.Enabled = false;
            btn_selectSavePath.Enabled = false;
            datePicker_startDate.Enabled = false;
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
            myDiseaseList.Initialize(AppDomain.CurrentDomain.BaseDirectory + @"Resource\2014ICD2217.xlsx", out Dic_DiseaseList);

            //DiseaseList myDiseaseList = new DiseaseList();
            //myDiseaseList.Initialize("", out List_Disease);

            //所有疾病的查询字典
            Dic myDic = new Dic();

            //进行统计
            processStatistics.Text = "统计开始";
            ServiceStatistics myStatistics = new ServiceStatistics();
            processStatistics.Text =  myStatistics.statistics(ref myDic, mnStatisticsParameters);
            //进行输出结果
            processOutputExcel.Text = "输出Excel开始";
            ServiceOutputExcel myOutputExcel = new ServiceOutputExcel();
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
