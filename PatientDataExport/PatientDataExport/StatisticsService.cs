using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//数据库
using PatientDataExport.Data;

namespace PatientDataExport
{
    public class StatisticsService
    {
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

        public void NotICDAdd(string diagnosisName, ref Dic myDic, string patientSex, string patientAgeRange)
        {
            //心肌缺血
            bool xinjiquexue = false;
            string ICDxinjiquexue = "ZNK003";
            if (diagnosisName.Contains("心肌缺血")) xinjiquexue = true;
            //心肌缺血
            if (xinjiquexue == true)
            {
                AddDic(ref myDic, ICDxinjiquexue, patientSex, patientAgeRange);
            }
        }

        //统计每个人的一种疾病的多个诊断指标
        public void MultiDiagnosis(List<string> eachPersonAllDiagnosis, ref Dic myDic, string patientSex, string patientAgeRange)
        {
            //高血压
            //收缩压高 F-31113 舒张压高 R03.002
            bool gaoxueya = false;
            string ICDgaoxueya = "ZNK001";
            //高脂血症  
            //高总胆固醇血症   R79.911
            //低密度脂蛋白增高  R79.916
            //高甘油三酯血症   R79.913
            bool gaozhixuezheng = false;
            string ICDgaozhixuezheng = "ZNK002";
            //白内障
            //老年性白内障    H25.901
            //白内障   H26.901
            bool baineizhang = false;
            string ICDbaineizhang = "ZWG001";
            //肝功能异常（实验室检查）
            //总胆红素增高    R79.901
            //直接胆红素增高   R79.902
            //简介胆红素增高   
            bool gangongnengyichang = false;
            string ICDgangongnengyichang = "ZSY001";
            //心律失常
            //房室传导阻滞 I44.302
            //房颤    R76.902
            //完全性右束支传导阻滞    I45.102
            //不完全性右束支传导阻滞 I45.101
            //左前分支传导阻滞 I44.401
            bool xinlvshichang = false;
            string ICDxinlvshichang = "ZNK006";
            foreach (string tempeachDiagnosis in eachPersonAllDiagnosis)
            {
                //高血压
                if (tempeachDiagnosis.Contains("R03.002") || tempeachDiagnosis.Contains("F-31113")) gaoxueya = true;
                //高脂血症
                if (tempeachDiagnosis.Contains("R79.911") || tempeachDiagnosis.Contains("R79.916") || tempeachDiagnosis.Contains("R79.913")) gaozhixuezheng = true;
                //白内障
                if (tempeachDiagnosis.Contains("H25.901") || tempeachDiagnosis.Contains("H26.901")) baineizhang = true;
                //肝功能异常（实验室检查）
                if (tempeachDiagnosis.Contains("R79.901") || tempeachDiagnosis.Contains("R79.902")) gangongnengyichang = true;
                //心律失常
                if (tempeachDiagnosis.Contains("I44.302") || tempeachDiagnosis.Contains("R76.902") || tempeachDiagnosis.Contains("I45.102") || tempeachDiagnosis.Contains("I45.101") || tempeachDiagnosis.Contains("I44.401") ) xinlvshichang = true;
            }
            //高血压
            if (gaoxueya == true)
            {
                AddDic(ref myDic,ICDgaoxueya, patientSex, patientAgeRange);
            }
            //高血脂
            if (gaozhixuezheng == true)
            {
                AddDic(ref myDic, ICDgaozhixuezheng, patientSex, patientAgeRange);
            }
            //白内障
            if (baineizhang == true)
            {
                AddDic(ref myDic, ICDbaineizhang, patientSex, patientAgeRange);
            }
            //肝功能异常（实验室检查）
            if (gangongnengyichang == true)
            {
                AddDic(ref myDic, ICDgangongnengyichang, patientSex, patientAgeRange);
            }
            //
            if (xinlvshichang == true)
            {
                AddDic(ref myDic, ICDxinlvshichang, patientSex, patientAgeRange);
            }
        }

        /// <summary>
        /// 一次向四个统计字典中增加统计项目
        /// </summary>
        /// <param name="myDic"></param>
        /// <param name="DiseaseICD"></param>
        /// <param name="patientSex"></param>
        /// <param name="patientAgeRange"></param>
        public void AddDic(ref Dic myDic, string DiseaseICD, string patientSex, string patientAgeRange)
        {
            AddDic_AllDisease(DiseaseICD, ref myDic);
            AddDic_Sex(DiseaseICD + "，" + patientSex, ref myDic);
            AddDic_Age(DiseaseICD + "，" + patientAgeRange, ref myDic);
            AddDic_SexAge(DiseaseICD + "，" + patientSex + "，" + patientAgeRange, ref myDic);
        }

        /// <summary>
        /// 统计所有的疾病
        /// </summary>
        /// <param name="diagcode"></param>
        /// <param name="myDic"></param>
        public void AddDic_AllDisease(string diagcode, ref Dic myDic)
        {
            //统计全体疾病
            if (myDic.Disease.ContainsKey(diagcode))
            {
                //已经遇到此种ICD疾病了
                myDic.Disease[diagcode]++;
            }
            else
            {
                //第一次统计此种ICD值疾病的数量
                myDic.Disease.Add(diagcode, 1);
            } 
        }

        /// <summary>
        /// 分年龄统计疾病发病率
        /// </summary>
        /// <param name="sep_DiseaseAge"></param>
        /// <param name="myDic"></param>
        public void AddDic_Age(string sep_DiseaseAge,ref Dic myDic)
        {
            //分年龄统计的
            if (myDic.DiseaseAge.ContainsKey(sep_DiseaseAge))
            {
                //已经遇到此种ICD疾病了
                myDic.DiseaseAge[sep_DiseaseAge]++;
            }
            else
            {
                //第一次统计此种ICD值疾病的数量
                myDic.DiseaseAge.Add(sep_DiseaseAge, 1);
            } 
        }
        /// <summary>
        /// 分性别分组
        /// </summary>
        /// <param name="sep_DiseaseSex"></param>
        /// <param name="myDic"></param>
        public void AddDic_Sex(string sep_DiseaseSex, ref Dic myDic)
        {
            //分性别
            if (myDic.DiseaseSex.ContainsKey(sep_DiseaseSex))
            {
                //已经有此性别分类病种
                myDic.DiseaseSex[sep_DiseaseSex]++;
            }
            else
            {
                //第一次统计此性别疾病
                myDic.DiseaseSex.Add(sep_DiseaseSex, 1);
            } 
        }
        /// <summary>
        /// 分年龄和性别分组
        /// </summary>
        /// <param name="sep_DiseaseSexAge"></param>
        /// <param name="myDic"></param>
        public void AddDic_SexAge(string sep_DiseaseSexAge, ref Dic myDic)
        {
            //分年龄和性别的
            if (myDic.DiseaseAgeSex.ContainsKey(sep_DiseaseSexAge))
            {
                //此年龄和性别分组
                myDic.DiseaseAgeSex[sep_DiseaseSexAge]++;
            }
            else
            {
                myDic.DiseaseAgeSex.Add(sep_DiseaseSexAge, 1);
            } 
        }

        //统计
        public string statistics(DateTime startDate, DateTime endDate, ref Dic myDic)
        {
            //全部疾病诊断列表名称
            Dictionary<string, string> List_Disease = new Dictionary<string, string>();

            DiseaseList myDiseaseList = new DiseaseList();
            myDiseaseList.Initialize("", out List_Disease);
            System.Windows.Forms.MessageBox.Show(List_Disease.First().ToString());

            //实际统计的人数，把年龄为零的人排除在外
            myDic.NumAll = 0;
            //查询数据库
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
            if (ExportResult == null)
            {
                return "没有查询到相应的患者";
            }


            //所有的性别分布范围
            //string eachPersonSex = "";
            //所有的年龄性别分布范围
            string eachPersonAgeSexRange = "";
            //每位患者的所有疾病
            List<string> tempPersonAllDisese = new List<string>();

            //遍历所有的患者
            foreach (var checkpatient in ExportResult)
            {
                if (checkpatient.age == null)
                {
                    continue;
                }
                //此人的年龄范围
                string patientAgeRange = AgeSeprate(checkpatient.age.ToString());
                //把年龄为零和空白的都排除在外
                if (patientAgeRange == "0" || patientAgeRange == "空白") continue;
                //真正进入统计的人数
                myDic.NumAll++;

                //区分每个人，使用Checkcode
                string eachPerson = checkpatient.checkcode.ToString();

                //统计不同年龄范围内的人群
                if (myDic.NumAge.ContainsKey(patientAgeRange))
                {
                    //将此年龄范围的人数加1
                    myDic.NumAge[patientAgeRange]++;
                }
                else
                {
                    //向统计的词典中增加此年龄范围
                    myDic.NumAge.Add(patientAgeRange, 1);
                }
                //每个病人的性别
                string patientSex = checkpatient.a0107.ToString();
                

                //统计不同性别的人群
                if (myDic.NumSex.ContainsKey(patientSex))
                {
                    //将此性别的人数加1
                    myDic.NumSex[patientSex]++;
                }
                else
                {
                    //第一次统计此性别
                    myDic.NumSex.Add(patientSex, 1);
                }

                //统计不同性别的年龄分布
                eachPersonAgeSexRange = patientSex + "，" + patientAgeRange;

                if (myDic.NumSexAge.ContainsKey(eachPersonAgeSexRange))
                {
                    myDic.NumSexAge[eachPersonAgeSexRange]++;
                }
                else
                {
                    myDic.NumSexAge.Add(eachPersonAgeSexRange, 1);
                }

                //所有的疾病
                try
                {
                    var diseaseResult = from s5 in myMedBaseEntities.hdatadiag where checkpatient.checkcode == s5.checkcode select s5;
                    //此病人检查无任何诊断
                    if (diseaseResult == null)
                    {
                        //跳过此人的循环
                        continue;
                    }
                    else
                    {
                        //此病人有诊断
                        foreach (var eachDisease in diseaseResult)
                        {
                            //存储此人的所有疾病
                            tempPersonAllDisese.Clear();

                            //不论有没有确定的ICD值，都要增加总的疾病数量
                            myDic.DiseaseNum++;
                            //诊断有确定ICD值，相应的疾病ICD值加1
                            if (eachDisease.diagcode != null)
                            {
                                //一次增加四个统计
                                AddDic(ref myDic, eachDisease.diagcode, patientSex, patientAgeRange);

                                //有确定的ICD值，那每个人的ICD确定诊断数目加1
                                myDic.ICDDiseaseNum++;
                            }
                            //诊断没有确定的ICD值
                            else
                            {
                                //没有确定的ICD值，那此人的ICD不确定诊断数目加1
                                myDic.NotICDDiseaseNum++;
                                NotICDAdd(eachDisease.diagname, ref myDic, patientSex, patientAgeRange);
                            }
                        }//循环每个人的所有疾病结束

                        //循环结束此人的所有诊断，再确定其中多项诊断确定一个疾病的诊断
                        MultiDiagnosis(tempPersonAllDisese,ref myDic, patientSex, patientAgeRange);

                    }//确定此人有诊断else结束
                }//try 查找此人的所有诊断结束
                catch
                {
                    System.Windows.Forms.MessageBox.Show("遍历查询到的患者时出现错误");
                }
                
            }//Foreach查询到的所有患者循环
            return "成功执行统计";
        }//结束public string statistics(DateTime startDate, DateTime endDate)
    }//结束public class StatisticsService
}
