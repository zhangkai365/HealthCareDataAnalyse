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
                               where s1.checkdate > startDate && s1.checkdate < endDate && (s1.a0704 != "01" && s1.a0704 != "02" && s1.a0704 != "03" && s1.a0704 != "04" && s1.a0704 != "05" && s1.a0704 != "14" && s1.a6405 == "02")
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
                                    string SexSep = Sex(checkpatient.a0107);
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
                    catch { }
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
                myWorkSheet.Cells[5, 7] = "总非ICD诊断的数量";
                int totalNotICDDiseaseNum = 0;
                foreach (var total_eachPersonNotICDDiseaseNum in Dic_DiseaseEachPersonNotICDNum)
                {
                    totalNotICDDiseaseNum = totalNotICDDiseaseNum + total_eachPersonNotICDDiseaseNum.Value;
                }
                myWorkSheet.Cells[5, 7] = totalNotICDDiseaseNum;

                //各个ICD诊断的数量
                int i = 0;
                myWorkSheet.Cells[6, 1] = "男，ICD诊断名称";
                myWorkSheet.Cells[6, 2] = "男，ICD诊断编码";
                myWorkSheet.Cells[6, 3] = "男，此ICD诊断发病数量";


                //筛选男性前20位的诊断
                try
                {
                    var top20male = (from temptop20all in Dic_DiseaseSex
                                    where temptop20all.Key.Contains("男，")
                                    orderby temptop20all.Value
                                    descending
                                    select temptop20all).Take(50);
                    if (top20male == null)
                    {
                        myWorkSheet.Cells[7, 2] = "空白";
                        myWorkSheet.Cells[7, 3] = "空白";
                    }
                    else
                    {
                        foreach (var eachICDDiseaseNum in top20male)
                        {
                            myWorkSheet.Cells[7 + i, 2] = eachICDDiseaseNum.Key.ToString();
                            myWorkSheet.Cells[7 + i, 3] = eachICDDiseaseNum.Value.ToString();
                            i++;
                        }
                    }
                }//此处筛选前20位诊断的查询过程，
                catch
                {
                    myWorkSheet.Cells[7, 2] = "异常";
                    myWorkSheet.Cells[7, 3] = "异常";
                }

                //筛选女性前20位的诊断
                try
                {
                    var top20female = (from temptop20all in Dic_DiseaseSex
                                     where temptop20all.Key.Contains("女，")
                                     orderby temptop20all.Value
                                     descending
                                     select temptop20all).Take(50);
                    if (top20female == null)
                    {
                        myWorkSheet.Cells[7, 45] = "空白";
                        myWorkSheet.Cells[7, 46] = "空白";
                    }
                    else
                    {
                        foreach (var eachICDDiseaseNum in top20female)
                        {
                            myWorkSheet.Cells[7 + i, 45] = eachICDDiseaseNum.Key.ToString();
                            myWorkSheet.Cells[7 + i, 46] = eachICDDiseaseNum.Value.ToString();
                            i++;
                        }
                    }
                }//此处筛选前20位诊断的查询过程，
                catch
                {
                    myWorkSheet.Cells[7, 2] = "异常";
                    myWorkSheet.Cells[7, 3] = "异常";
                }
                //分年龄、性别ICD诊断的数量
                //<45
                myWorkSheet.Cells[6, 5] = "男，<45 ICD诊断名称";
                myWorkSheet.Cells[6, 6] = "男，<45 ICD诊断编码";
                myWorkSheet.Cells[6, 7] = "男，<45 此ICD诊断发病数量";
                int templower45m = 0;
                //45-50
                myWorkSheet.Cells[6, 9] = "男，45-50 ICD诊断名称";
                myWorkSheet.Cells[6, 10] = "男，45-50 ICD诊断编码";
                myWorkSheet.Cells[6, 11] = "男，45-50 此ICD诊断发病数量";
                int temp45to50m = 0;
                //50-55
                myWorkSheet.Cells[6, 13] = "男，50-55 ICD诊断名称";
                myWorkSheet.Cells[6, 14] = "男，50-55 ICD诊断编码";
                myWorkSheet.Cells[6, 15] = "男，50-55 此ICD诊断发病数量";
                int temp50to55m = 0;
                //55-60
                myWorkSheet.Cells[6, 17] = "男，55-60 ICD诊断名称";
                myWorkSheet.Cells[6, 18] = "男，55-60 ICD诊断编码";
                myWorkSheet.Cells[6, 19] = "男，55-60 此ICD诊断发病数量";
                int temp55to60m = 0;
                //>60
                myWorkSheet.Cells[6, 21] = "男，>60 ICD诊断名称";
                myWorkSheet.Cells[6, 22] = "男，>60 ICD诊断编码";
                myWorkSheet.Cells[6, 23] = "男，>60 此ICD诊断发病数量";
                int temphigher60m = 0;
                //<45
                myWorkSheet.Cells[6, 25] = "女，<45 ICD诊断名称";
                myWorkSheet.Cells[6, 26] = "女，<45 ICD诊断编码";
                myWorkSheet.Cells[6, 27] = "女，<45 此ICD诊断发病数量";
                int templower45f = 0;
                //45-50
                myWorkSheet.Cells[6, 29] = "女，45-50 ICD诊断名称";
                myWorkSheet.Cells[6, 30] = "女，45-50 ICD诊断编码";
                myWorkSheet.Cells[6, 31] = "女，45-50 此ICD诊断发病数量";
                int temp45to50f = 0;
                //50-55
                myWorkSheet.Cells[6, 33] = "女，50-55 ICD诊断名称";
                myWorkSheet.Cells[6, 34] = "女，50-55 ICD诊断编码";
                myWorkSheet.Cells[6, 35] = "女，50-55 此ICD诊断发病数量";
                int temp50to55f = 0;
                //55-60
                myWorkSheet.Cells[6, 37] = "女，55-60 ICD诊断名称";
                myWorkSheet.Cells[6, 38] = "女，55-60 ICD诊断编码";
                myWorkSheet.Cells[6, 39] = "女，55-60 此ICD诊断发病数量";
                int temp55to60f = 0;
                //>60
                myWorkSheet.Cells[6, 41] = "女，>60 ICD诊断名称";
                myWorkSheet.Cells[6, 42] = "女，>60 ICD诊断编码";
                myWorkSheet.Cells[6, 43] = "女，>60 此ICD诊断发病数量";
                int temphigher60f = 0;



                //筛选<45的前20位诊断
                try
                {
                    var top20lower45m = (from temptop20lower45 in Dic_DiseaseAgeSex
                                        where temptop20lower45.Key.Contains("男，<45")
                                        orderby temptop20lower45.Value
                                        descending
                                        select temptop20lower45).Take(20);
                    if (top20lower45m == null)
                    {
                        myWorkSheet.Cells[7, 6] = "空白";
                        myWorkSheet.Cells[7, 7] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top20lower45m)
                        {
                            myWorkSheet.Cells[7 + templower45m, 6] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + templower45m, 7] = eachDiagnosis.Value.ToString();
                            templower45m++;
                        }
                    }
                }//筛选<45岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[8, 6] = "异常";
                    myWorkSheet.Cells[8, 7] = "异常"; 
                }

                //筛选45-50前20位诊断
                try
                {
                    var top2045to50m = (from temptop2045to50 in Dic_DiseaseAgeSex
                                        where temptop2045to50.Key.Contains("男，45-50")
                                        orderby temptop2045to50.Value
                                        descending
                                        select temptop2045to50).Take(20);
                    if (top2045to50m == null)
                    {
                        myWorkSheet.Cells[7, 10] = "空白";
                        myWorkSheet.Cells[7, 11] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top2045to50m)
                        {
                            myWorkSheet.Cells[7 + temp45to50m, 10] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + temp45to50m, 11] = eachDiagnosis.Value.ToString();
                            temp45to50m++;
                        }
                    }
                }//筛选45-50岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[7, 10] = "异常";
                    myWorkSheet.Cells[7, 11] = "异常";
                }

                //筛选50-55前20位诊断
                try
                {
                    var top2050to55m = (from temptop2050to55 in Dic_DiseaseAgeSex
                                       where temptop2050to55.Key.Contains("男，50-55")
                                       orderby temptop2050to55.Value
                                       descending
                                       select temptop2050to55).Take(20);
                    if (top2050to55m == null)
                    {
                        myWorkSheet.Cells[7, 14] = "空白";
                        myWorkSheet.Cells[7, 15] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top2050to55m)
                        {
                            myWorkSheet.Cells[7 + temp50to55m, 14] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + temp50to55m, 15] = eachDiagnosis.Value.ToString();
                            temp50to55m++;
                        }
                    }
                }//筛选50-55岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[7, 14] = "异常";
                    myWorkSheet.Cells[7, 15] = "异常";
                }

                //筛选55-60前20位诊断
                try
                {
                    var top2055to60m = (from temptop2055to60 in Dic_DiseaseAgeSex
                                       where temptop2055to60.Key.Contains("男，55-60")
                                       orderby temptop2055to60.Value
                                       descending
                                       select temptop2055to60).Take(20);
                    if (top2055to60m == null)
                    {
                        myWorkSheet.Cells[7, 18] = "空白";
                        myWorkSheet.Cells[7, 19] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top2055to60m)
                        {
                            myWorkSheet.Cells[7 + temp55to60m, 18] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + temp55to60m, 19] = eachDiagnosis.Value.ToString();
                            temp55to60m++;
                        }
                    }
                }//筛选50-55岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[7, 18] = "异常";
                    myWorkSheet.Cells[7, 19] = "异常";
                }

                //筛选>60的前20位诊断
                try
                {
                    var top20higher60m = (from temptop20higher60 in Dic_DiseaseAgeSex
                                        where temptop20higher60.Key.Contains("男，>60")
                                        orderby temptop20higher60.Value
                                        descending
                                        select temptop20higher60).Take(20);
                    if (top20higher60m == null)
                    {
                        myWorkSheet.Cells[7, 22] = "空白";
                        myWorkSheet.Cells[7, 23] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top20higher60m)
                        {
                            myWorkSheet.Cells[7 + temphigher60m, 22] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + temphigher60m, 23] = eachDiagnosis.Value.ToString();
                            temphigher60m++;
                        }
                    }
                }//筛选>60岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[8, 22] = "异常";
                    myWorkSheet.Cells[8, 23] = "异常";
                }



                //筛选<45的前20位诊断
                try
                {
                    var top20lower45f = (from temptop20lower45 in Dic_DiseaseAgeSex
                                        where temptop20lower45.Key.Contains("女，<45")
                                        orderby temptop20lower45.Value
                                        descending
                                        select temptop20lower45).Take(20);
                    if (top20lower45f == null)
                    {
                        myWorkSheet.Cells[7, 26] = "空白";
                        myWorkSheet.Cells[7, 27] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top20lower45f)
                        {
                            myWorkSheet.Cells[7 + templower45f, 26] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + templower45f, 27] = eachDiagnosis.Value.ToString();
                            templower45f++;
                        }
                    }
                }//筛选<45岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[7, 26] = "异常";
                    myWorkSheet.Cells[7, 27] = "异常";
                }

                //筛选45-50前20位诊断
                try
                {
                    var top2045to50f = (from temptop2045to50 in Dic_DiseaseAgeSex
                                       where temptop2045to50.Key.Contains("女，45-50")
                                       orderby temptop2045to50.Value
                                       descending
                                       select temptop2045to50).Take(20);
                    if (top2045to50f == null)
                    {
                        myWorkSheet.Cells[7, 30] = "空白";
                        myWorkSheet.Cells[7, 31] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top2045to50f)
                        {
                            myWorkSheet.Cells[7 + temp45to50f, 30] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + temp45to50f, 31] = eachDiagnosis.Value.ToString();
                            temp45to50f++;
                        }
                    }
                }//筛选45-50岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[7, 30] = "异常";
                    myWorkSheet.Cells[7, 31] = "异常";
                }

                //筛选50-55前20位诊断
                try
                {
                    var top2050to55f = (from temptop2050to55 in Dic_DiseaseAgeSex
                                       where temptop2050to55.Key.Contains("女，50-55")
                                       orderby temptop2050to55.Value
                                       descending
                                       select temptop2050to55).Take(20);
                    if (top2050to55f == null)
                    {
                        myWorkSheet.Cells[7, 34] = "空白";
                        myWorkSheet.Cells[7, 35] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top2050to55f)
                        {
                            myWorkSheet.Cells[7 + temp50to55f, 34] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + temp50to55f, 35] = eachDiagnosis.Value.ToString();
                            temp50to55f++;
                        }
                    }
                }//筛选50-55岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[7, 34] = "异常";
                    myWorkSheet.Cells[7, 35] = "异常";
                }

                //筛选55-60前20位诊断
                try
                {
                    var top2055to60f = (from temptop2055to60 in Dic_DiseaseAgeSex
                                       where temptop2055to60.Key.Contains("女，55-60")
                                       orderby temptop2055to60.Value
                                       descending
                                       select temptop2055to60).Take(20);
                    if (top2055to60f == null)
                    {
                        myWorkSheet.Cells[7, 38] = "空白";
                        myWorkSheet.Cells[7, 39] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top2055to60f)
                        {
                            myWorkSheet.Cells[7 + temp55to60f, 38] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + temp55to60f, 39] = eachDiagnosis.Value.ToString();
                            temp55to60f++;
                        }
                    }
                }//筛选50-55岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[7, 38] = "异常";
                    myWorkSheet.Cells[7, 39] = "异常";
                }

                //筛选>60的前20位诊断
                try
                {
                    var top20higher60f = (from temptop20higher60 in Dic_DiseaseAgeSex
                                         where temptop20higher60.Key.Contains("女，>60")
                                         orderby temptop20higher60.Value
                                         descending
                                         select temptop20higher60).Take(20);
                    if (top20higher60f == null)
                    {
                        myWorkSheet.Cells[7, 42] = "空白";
                        myWorkSheet.Cells[7, 43] = "空白";
                    }
                    else
                    {
                        foreach (var eachDiagnosis in top20higher60f)
                        {
                            myWorkSheet.Cells[7 + temphigher60f, 42] = eachDiagnosis.Key.ToString();
                            myWorkSheet.Cells[7 + temphigher60f, 43] = eachDiagnosis.Value.ToString();
                            temphigher60f++;
                        }
                    }
                }//筛选>60岁的前20位诊断的查询
                catch
                {
                    myWorkSheet.Cells[7, 42] = "异常";
                    myWorkSheet.Cells[7, 43] = "异常";
                }

                //保存文件
                myWorkbook.SaveAs(FilePath);
                myWorkbook.Close();
                myExcel.Quit();
                iffinished.Text = "已完成！";
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

    }
}
