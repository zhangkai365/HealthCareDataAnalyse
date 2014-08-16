using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

//include
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
            //设置ICD存储文件的路径及文件名
            string ICDFileNameOpen = @"C:\Users\Win7x64_20140606\Desktop\2014年查体工作总结（修改201408080753）\疾病总表\ICD1.xls";
            string ICDFileNameStore = @"C:\Users\Win7x64_20140606\Desktop\2014年查体工作总结（修改201408080753）\疾病总表\ICD2.xls";
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

            //打开所有诊断的总表
            Excel.Workbook allICDWorkbook = myExcel.Workbooks.Open(ICDFileNameOpen);
            Excel.Worksheet allICDWorksheet = allICDWorkbook.Worksheets[1];

            //全部疾病诊断列表名称
            Dictionary<string, string> List_Disease = new Dictionary<string, string>();
            int count_empty = 0;
            for (int i = 0; i < 100; i++)
            {
                try
                {
                    List_Disease.Add(ICDWorksheet.get_Range("B" + i).Text, ICDWorksheet.get_Range("A" + i).Text);
                }
                catch
                {
                    count_empty++;
                    continue;
                }
                if (count_empty > 2) break;
            }

            //统计不同年龄段的人数
            Dictionary<string, int> Dic_NumAge = new Dictionary<string, int>();
            //统计不同性别的人数
            Dictionary<string, int> Dic_NumSex = new Dictionary<string, int>();
            //统计不同年龄段和不同性别的人数
            Dictionary<string, int> Dic_NumSexAge = new Dictionary<string, int>(); 

            //疾病统计
            Dictionary<string, int> Dic_Disease = new Dictionary<string, int>();
            //分性别统计
            Dictionary<string, int> Dic_DiseaseSex = new Dictionary<string, int>();
            //分年龄统计 疾病统计
            Dictionary<string,int> Dic_DiseaseAge = new Dictionary<string,int>();
            //分年龄性别统计
            Dictionary<string, int> Dic_DiseaseAgeSex = new Dictionary<string, int>();
            //每个人的诊断数目统计
            Dictionary<string,int> Dic_DiseaseEachPersonNum = new Dictionary<string,int>();
            //每个人的有确定ICD10疾病数目
            Dictionary<string,int> Dic_DiseaseEachPersonICDNum = new Dictionary<string,int>();
            //每个人无确定ICD诊断数目
            Dictionary<string,int> Dic_DiseaseEachPersonNotICDNum = new Dictionary<string,int>();

            //进行处理的人数
            int peoplecount = 0;

            medbaseEntities myMedBaseEntities = new medbaseEntities();
            //查询所有的待查询时间段内检查的患者
            //查询条件  a0704 任职级别 01 副市级 02 正局级 03 副局级 04 正高 05 副高 14 院士
            //查询条件  a6405 在职情况
            //全部包括  (s1.a0704 == "01" || s1.a0704 == "02" || s1.a0704 == "03" || s1.a0704 == "04" || s1.a0704 == "05" || s1.a0704 == "14" || s1.a6405 == "02")
            //副市级  s1.a0704 == "01"
            //正局级  s1.a0704 == "02"
            //副局级  s1.a0704 == "03"
            //高级知识分子  (s1.a0704 == "04" || s1.a0704 == "05" || s1.a0704 == "14")
            //离休  s1.a6405 == "02"
            //离休 解决与上面重复问题 (s1.a0704 != "01" && s1.a0704 != "02" && s1.a0704 != "03" && s1.a0704 != "04" && s1.a0704 != "05" && s1.a0704 != "14" && s1.a6405 == "02")
            var ExportResult = from s1 in myMedBaseEntities.hcheckmemb
                               where s1.checkdate > startDate && s1.checkdate < endDate && (s1.a0704 == "01" || s1.a0704 == "02" || s1.a0704 == "03" || s1.a0704 == "04" || s1.a0704 == "05" || s1.a0704 == "14" || s1.a6405 == "02")
                               select s1;
            //总数
            totalNum.Text = ExportResult.Count().ToString();
            //确保病人数量不为空
            if (ExportResult == null)
            {
                MessageBox.Show("搜索结果为空，程序退出。");
                return;
            }
            else
            {
                //输出总人数
                myWorkSheet.Cells[1, 1] = "总人数";
                myWorkSheet.Cells[1, 2] = ExportResult.Count();
                
                //所有的年龄分布范围
                string eachPersonAgeRange = "";
                //所有的性别分布范围
                string eachPersonSex = "";
                //所有的年龄性别分布范围
                string eachPersonAgeSexRange = "";
                //存储性别
                string tempSex = "";
                
                //遍历所有的患者
                foreach (var checkpatient in ExportResult)
                {
                    //统计不同年龄段的人数
                    //此人的年龄范围
                    if (checkpatient.age == null)
                    {
                        continue;
                    }
                    eachPersonAgeRange = AgeSeprate(checkpatient.age.ToString());
                    //把年龄为零和空白的都排除在外
                    if (eachPersonAgeRange == "0" || eachPersonAgeRange == "空白") continue;
                    //真正进入统计的人数
                    peoplecount++;
                    progressNum.Text = peoplecount.ToString();

                    //区分每个人，使用Checkcode
                    string eachPerson = checkpatient.checkcode.ToString();

                    //统计不同年龄范围内的人群
                    if (Dic_NumAge.ContainsKey(eachPersonAgeRange))
                    {
                        //将此年龄范围的人数加1
                        Dic_NumAge[eachPersonAgeRange]++;
                    }
                    else
                    {
                        //向统计的词典中增加此年龄范围
                        Dic_NumAge.Add(eachPersonAgeRange, 1);
                    }

                    tempSex = Sex(checkpatient.a0107);

                    //统计不同性别的人群
                    if (Dic_NumSex.ContainsKey(tempSex))
                    {
                        //将此性别的人数加1
                        Dic_NumSex[tempSex]++;
                    }
                    else
                    {
                        //第一次统计此性别
                        Dic_NumSex.Add(tempSex, 1);
                    }

                    //统计不同性别的年龄分布
                    eachPersonAgeSexRange = tempSex + "，" + eachPersonAgeRange;

                    if (Dic_NumSexAge.ContainsKey(eachPersonAgeSexRange))
                    {
                        Dic_NumSexAge[eachPersonAgeSexRange]++;
                    }
                    else
                    {
                        Dic_NumSexAge.Add(eachPersonAgeSexRange, 1);
                    }

                    //所有的疾病
                    try
                    {
                        var diseaseResult = from s5 in myMedBaseEntities.hdatadiag where checkpatient.checkcode == s5.checkcode select s5;
                        //此病人检查无任何诊断
                        if (diseaseResult == null)
                        {
                            //每个病人的疾病总数为0
                            Dic_DiseaseEachPersonNum.Add(eachPerson, 0);
                            //每个病人的ICD确定疾病数为0
                            Dic_DiseaseEachPersonICDNum.Add(eachPerson, 0);
                            //每个病人的无ICD确定值疾病数为0
                            Dic_DiseaseEachPersonNotICDNum.Add(eachPerson, 0);
                            //跳过此人的循环
                            continue;
                        }
                        else
                        {
                            //此病人有诊断
                            foreach (var eachDisease in diseaseResult)
                            {
                                //不论有没有确定的ICD值，都要增加总的疾病数量
                                if (Dic_DiseaseEachPersonNum.ContainsKey(eachPerson))
                                {
                                    Dic_DiseaseEachPersonNum[eachPerson]++;
                                }
                                else
                                {
                                    //第一次统计此病人的疾病总量
                                    Dic_DiseaseEachPersonNum.Add(eachPerson, 1);
                                }
                                //诊断有确定ICD值，相应的疾病ICD值加1
                                if (eachDisease.diagcode != null)
                                {
                                    //区分患者的年龄
                                    string ageSep = AgeSeprate(checkpatient.age.ToString()) + "，" + eachDisease.diagcode.ToString();
                                    //区分患者的性别
                                    string SexSep = Sex(checkpatient.a0107) + "，" + eachDisease.diagcode.ToString();
                                    //区分患者的年龄和性别
                                    string ageSexSep = Sex(checkpatient.a0107) + "，" + ageSep;
                                    //不分年龄的
                                    if (Dic_Disease.ContainsKey(eachDisease.diagcode.ToString()))
                                    {
                                        //已经遇到此种ICD疾病了
                                        //不分年龄
                                        Dic_Disease[eachDisease.diagcode]++;
                                    }
                                    else
                                    {
                                        //第一次统计此种ICD值疾病的数量
                                        //不分年龄
                                        Dic_Disease.Add(eachDisease.diagcode, 1);
                                    }
                                    //分年龄的
                                    if (Dic_DiseaseAge.ContainsKey(ageSep))
                                    {
                                        //分年龄
                                        //已经遇到此种ICD疾病了
                                        Dic_DiseaseAge[ageSep]++;
                                    }
                                    else
                                    {
                                        //分年龄
                                        //第一次统计此种ICD值疾病的数量
                                        Dic_DiseaseAge.Add(ageSep, 1);
                                    }

                                    //分性别
                                    if (Dic_DiseaseSex.ContainsKey(SexSep))
                                    {
                                        //已经有此性别分类病种
                                        Dic_DiseaseSex[SexSep]++;
                                    }
                                    else
                                    {
                                        //第一次统计此性别疾病
                                        Dic_DiseaseSex.Add(SexSep, 1);
                                    }
                                    //分年龄和性别的
                                    if (Dic_DiseaseAgeSex.ContainsKey(ageSexSep))
                                    {
                                        //此年龄和性别分组
                                        Dic_DiseaseAgeSex[ageSexSep]++;
                                    }
                                    else
                                    {
                                        Dic_DiseaseAgeSex.Add(ageSexSep, 1);
                                    }

                                    //有确定的ICD值，那每个人的ICD确定诊断数目加1
                                    if (Dic_DiseaseEachPersonICDNum.ContainsKey(eachPerson))
                                    {
                                        Dic_DiseaseEachPersonICDNum[eachPerson]++;
                                    }
                                    else
                                    {
                                        //第一次统计此人的ICD值
                                        Dic_DiseaseEachPersonICDNum.Add(eachPerson, 1);
                                    }
                                }
                                //诊断没有确定的ICD值
                                else
                                {
                                    //没有确定的ICD值，那此人的ICD不确定诊断数目加1
                                    if (Dic_DiseaseEachPersonNotICDNum.ContainsKey(eachPerson))
                                    {
                                        Dic_DiseaseEachPersonNotICDNum[eachPerson]++;
                                    }
                                    else
                                    {
                                        //第一次统计此人的非ICD值
                                        Dic_DiseaseEachPersonNotICDNum.Add(eachPerson, 1);
                                    }
                                }
                            }
                        }
                    }//确定此病人有诊断的结束
                    catch 
                    { }
                }


                //输出统计结果
                //输出处理的人数
                myWorkSheet.Cells[1, 3] = "统计的人数";
                myWorkSheet.Cells[1, 4] = peoplecount;
                //不同性别的人数
                myWorkSheet.Cells[1, 5] = "男";
                myWorkSheet.Cells[1, 6] = Dic_NumSex["男"];

                myWorkSheet.Cells[1, 7] = "女";
                myWorkSheet.Cells[1, 8] = Dic_NumSex["女"];
                //不同年龄段的人数范围
                int tempCountRangeNum = 0;
                foreach (var eachRangeNum in Dic_NumAge)
                {
                    myWorkSheet.Cells[2, tempCountRangeNum + 1] = eachRangeNum.Key.ToString();
                    myWorkSheet.Cells[2, tempCountRangeNum + 2] = eachRangeNum.Value.ToString();
                    tempCountRangeNum = tempCountRangeNum + 2;
                }
                //不同年龄段男女数目
                int tempCountSexAgeNum = 0;
                foreach (var eachSexAgeNum in Dic_NumSexAge)
                {
                    myWorkSheet.Cells[3, tempCountSexAgeNum + 1] = eachSexAgeNum.Key.ToString();
                    myWorkSheet.Cells[3, tempCountSexAgeNum + 2] = eachSexAgeNum.Value.ToString();
                    tempCountSexAgeNum = tempCountSexAgeNum + 2;
                }

                //总疾病的数量
                myWorkSheet.Cells[5, 1] = "总疾病的数量";
                int totalDiseaseNum = 0;
                foreach (var total_eachPersonDiseaseNum in Dic_DiseaseEachPersonNum)
                {
                    totalDiseaseNum = totalDiseaseNum + total_eachPersonDiseaseNum.Value;
                }
                myWorkSheet.Cells[5, 2] = totalDiseaseNum;

                //总ICD诊断数量
                myWorkSheet.Cells[5, 4] = "总ICD诊断的数量";
                int totalICDDiseaseNum = 0;
                foreach (var total_eachPersonICDDiseaseNum in Dic_DiseaseEachPersonICDNum)
                {
                    totalICDDiseaseNum = totalICDDiseaseNum + total_eachPersonICDDiseaseNum.Value;
                }
                myWorkSheet.Cells[5, 5] = totalICDDiseaseNum;
                
                //总非ICD诊断数量
                myWorkSheet.Cells[5, 6] = "总非ICD诊断的数量";
                int totalNotICDDiseaseNum = 0;
                foreach (var total_eachPersonNotICDDiseaseNum in Dic_DiseaseEachPersonNotICDNum)
                {
                    totalNotICDDiseaseNum = totalNotICDDiseaseNum + total_eachPersonNotICDDiseaseNum.Value;
                }
                myWorkSheet.Cells[5, 7] = totalNotICDDiseaseNum;

                //各个ICD诊断的数量
                //简化ICD表头
                //性别字符串集合
                string[] Collection_Sex = { "男", "女" };
                //疾病诊断
                string[] titleDiseasePart = { "ICD诊断名称", "ICD诊断编码", "此ICD诊断发病数量" };
                int titleRangeAll = 1;
                foreach (string preTitleSex in Collection_Sex)
                {
                    foreach (string titleDisease in titleDiseasePart)
                    {
                        myWorkSheet.Cells[6, titleRangeAll++] = preTitleSex + "," + titleDisease;
                    }
                }
                //筛选此范围内前50位的诊断，分男女
                //控制横向移动的变量
                int countX_Top20all = 0;
                foreach (string top20allSex in Collection_Sex)
                {
                    //控制纵向移动的变量
                    int countY_Top20all = 0;
                    try
                    {
                        var top20male = (from temptop20all in Dic_DiseaseSex
                                         where temptop20all.Key.Contains(top20allSex + "，")
                                         orderby temptop20all.Value
                                         descending
                                         select temptop20all).Take(100);
                        if (top20male == null)
                        {
                            myWorkSheet.Cells[7, 2 + countX_Top20all] = "空白";
                            myWorkSheet.Cells[7, 3 + countX_Top20all] = "空白";
                            ICDWorksheet.Cells[1, 2 + countX_Top20all] = "空白";
                            ICDWorksheet.Cells[1, 3 + countX_Top20all] = "空白";
                        }
                        else
                        {
                            foreach (var eachICDDiseaseNum in top20male)
                            {
                                //做出统计
                                myWorkSheet.Cells[7 + countY_Top20all, 1 + countX_Top20all] = DiagnosisName(List_Disease,eachICDDiseaseNum.Key);
                                myWorkSheet.Cells[7 + countY_Top20all, 2 + countX_Top20all] = eachICDDiseaseNum.Key.ToString();
                                myWorkSheet.Cells[7 + countY_Top20all, 3 + countX_Top20all] = eachICDDiseaseNum.Value.ToString();
                                //写入所有的诊断列表
                                ICDWorksheet.Cells[1 + countY_Top20all, 2 + countX_Top20all] = eachICDDiseaseNum.Key.ToString();
                                ICDWorksheet.Cells[1 + countY_Top20all, 3 +countX_Top20all] = eachICDDiseaseNum.Value.ToString();
                                //纵向移动
                                countY_Top20all++;
                            }
                        }
                    }
                    catch
                    {
                        myWorkSheet.Cells[7, 2] = "异常";
                        myWorkSheet.Cells[7, 3] = "异常";
                        ICDWorksheet.Cells[1, 2] = "异常";
                        ICDWorksheet.Cells[1, 3] = "异常";
                    }
                    //横向移动
                    countX_Top20all = countX_Top20all +3;
                }

                //分年龄、性别ICD诊断的数量的标题，简化
                string[] titleAge = { "<45", "45-50", "50-55", ">60" };
                //控制表头横向移动的变量
                int countX_titleSexAgeDisease = 0;
                foreach (string title_Sex in Collection_Sex)
                {
                    foreach (string title_Age in titleAge)
                    {
                        foreach (string title_Disease in titleDiseasePart)
                        {
                            myWorkSheet.Cells[6, 7 + countX_titleSexAgeDisease++] = title_Sex + "，" + title_Age + "，" + title_Disease ;
                        }
                    }
                }

                //控制疾病列表的横向移动
                int countX2_SexAgeDisease = 0;

                //循环性别
                foreach (string temp_titleSex in Collection_Sex)
                {
                    //循环年龄
                    foreach (string temp_titleAge in titleAge)
                    {
                        //控制疾病移动的纵向移动,每次循环之内的疾病都清零
                        int countY2_SexAgeDisease = 0;
                        try
                        {
                            var top20DiseaseOfEachAge = (from tempTop20Disease in Dic_DiseaseAgeSex
                                                         where tempTop20Disease.Key.Contains(temp_titleSex + "，" + temp_titleAge)
                                                         orderby tempTop20Disease.Value
                                                         descending
                                                         select tempTop20Disease).Take(50);
                            //此性别、年龄的患者的统计到疾病为空白
                            if (top20DiseaseOfEachAge == null)
                            {
                                //诊断名称
                                myWorkSheet.Cells[7 + countY2_SexAgeDisease, 7 + countX2_SexAgeDisease] = "空白";
                                //诊断编码
                                myWorkSheet.Cells[7 + countY2_SexAgeDisease, 8 + countX2_SexAgeDisease] = "空白";
                                //诊断出现的数量
                                myWorkSheet.Cells[7 + countY2_SexAgeDisease, 9 + countX2_SexAgeDisease] = "空白";
                            }
                            else
                            {
                                foreach (var eachDiseaseOfTop20DiseaseOfEachAge in top20DiseaseOfEachAge)
                                {
                                    //通过列表之中查找响应的疾病名称
                                    myWorkSheet.Cells[7 + countY2_SexAgeDisease, 7 + countX2_SexAgeDisease] = DiagnosisName(List_Disease, eachDiseaseOfTop20DiseaseOfEachAge.Key.ToString());
                                    //疾病的编码
                                    myWorkSheet.Cells[7 + countY2_SexAgeDisease, 8 + countX2_SexAgeDisease] = eachDiseaseOfTop20DiseaseOfEachAge.Key.ToString();
                                    //疾病的数量
                                    myWorkSheet.Cells[7 + countY2_SexAgeDisease, 9 + countX2_SexAgeDisease] = eachDiseaseOfTop20DiseaseOfEachAge.Value.ToString();
                                    countY2_SexAgeDisease++;
                                }
                            }
                        }
                        catch 
                        {
                            //诊断名称
                            myWorkSheet.Cells[7 + countY2_SexAgeDisease, 7 + countX2_SexAgeDisease] = "查询过程出现异常";
                            //诊断编码
                            myWorkSheet.Cells[7 + countY2_SexAgeDisease, 8 + countX2_SexAgeDisease] = "查询过程出现异常";
                            //诊断出现的数量
                            myWorkSheet.Cells[7 + countY2_SexAgeDisease, 9 + countX2_SexAgeDisease] = "查询过程出现异常";
                        }
                        //内层循环，每个年龄性别循环结束后，横向移动3个单元格
                        countX2_SexAgeDisease = countX2_SexAgeDisease + 3;
                    }
                }



                //保存文件
                myWorkbook.SaveAs(FilePath);
                myWorkbook.Close();
                //如果界面上面的创建新疾病列表的选项为选中，则保存新的疾病列表
                if (chk_CreateNewDiseaseList.Checked == true)
                {
                    ICDWorkbook.SaveAs(ICDFileNameStore);
                    ICDWorkbook.Close(); 
                }
                myExcel.Quit();
                iffinished.Text = @"已完成！";
            }//这里是确定搜索固定范围内的病人结果不为空

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
        //年龄分组
        public string AgeSeprate(string age)
        {
            short _age = 0;
            if (Int16.TryParse(age, out _age))
            {
                if (_age == 0) return "0"; 
                if (_age < 45) return "<45";
                if (_age >= 45 && _age < 50) return "45-50";
                if (_age >= 50 && _age < 55) return "50-55";
                if (_age >= 55 && _age < 60) return "55-60";
                if (_age >= 60) return ">60";
            }
            else
            {
                return "空白";
            }
            return "空白";
        }
        //性别分组
        public string Sex(string sex)
        {
            if (sex == "男") return "男";
            if (sex == "女") return "女";
            MessageBox.Show("性别定义出错");
            return "错误";
        }

        public void OutputCellContent(Excel.Worksheet outputWorksheet, int cellx, int celly, string cellContent)
        {
            outputWorksheet.Cells[cellx, celly] = cellContent;
        }
        //保留 用于选择连接的数据库
        public System.Data.SqlClient.SqlConnectionStringBuilder ConnectionString()
        {
            //<add name="medbaseEntities" 
            //connectionString="metadata=res://*/Data.PatientData.csdl|res://*/Data.PatientData.ssdl|res://*/Data.PatientData.msl;
            //provider=System.Data.SqlClient;provider connection string=&quot;
            //data source=192.168.1.161;
            //initial catalog=medbase;
            //user id=sa;password=@Zhangkai851983;
            //MultipleActiveResultSets=True;
            //App=EntityFramework&quot;" 
            //providerName="System.Data.EntityClient" />
            System.Data.SqlClient.SqlConnectionStringBuilder myConnectionString = new System.Data.SqlClient.SqlConnectionStringBuilder();
            myConnectionString.ConnectionString = @"metadata=res://*/Data.PatientData.csdl|res://*/Data.PatientData.ssdl|res://*/Data.PatientData.msl;provider=System.Data.SqlClient;provider connection string=&quot;data source=192.168.1.161;initial catalog=medbase;user id=sa;password=@Zhangkai851983;MultipleActiveResultSets=True;App=EntityFramework&quot;";
            return myConnectionString;
        }

        /// <summary>
        /// 确定疾病的名称
        /// </summary>
        /// <param name="allICDDisease">全部疾病的列表</param>
        /// <param name="ICDCode">待确定的疾病编码</param>
        /// <returns></returns>
        public string DiagnosisName(Dictionary<string,string> allICDDisease ,string ICDCode)
        {
            if (allICDDisease.ContainsKey(ICDCode))
            {
                return allICDDisease[ICDCode].ToString();
            }
            return "无此编码疾病信息";
        }

    }
}
