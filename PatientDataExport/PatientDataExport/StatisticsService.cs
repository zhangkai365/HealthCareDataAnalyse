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
                               where s1.checkdate > startDate && s1.checkdate < endDate && (s1.a0704 == "01" || s1.a0704 == "02" || s1.a0704 == "03" || s1.a0704 == "04" || s1.a0704 == "05" || s1.a0704 == "14" || s1.a6405 == "02")
                               select s1;
            if (ExportResult == null)
            {
                return "没有查询到相应的患者";
            }


            //所有的性别分布范围
            string eachPersonSex = "";
            //所有的年龄性别分布范围
            string eachPersonAgeSexRange = "";


            //遍历所有的患者
            foreach (var checkpatient in ExportResult)
            {
                //患者的性别

                //患者的年龄分组

                //统计不同年龄段的人数

                if (checkpatient.age == null)
                {
                    continue;
                }
                //此人的年龄范围
                string PatatienAgeRange = AgeSeprate(checkpatient.age.ToString());
                //把年龄为零和空白的都排除在外
                if (PatatienAgeRange == "0" || PatatienAgeRange == "空白") continue;
                //真正进入统计的人数
                myDic.NumAll++;

                //区分每个人，使用Checkcode
                string eachPerson = checkpatient.checkcode.ToString();

                //统计不同年龄范围内的人群
                if (myDic.NumAge.ContainsKey(PatatienAgeRange))
                {
                    //将此年龄范围的人数加1
                    myDic.NumAge[PatatienAgeRange]++;
                }
                else
                {
                    //向统计的词典中增加此年龄范围
                    myDic.NumAge.Add(PatatienAgeRange, 1);
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
                eachPersonAgeSexRange = patientSex + "，" + PatatienAgeRange;

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
                        //每个病人的疾病总数为0
                        myDic.DiseaseEachPersonNum.Add(eachPerson, 0);
                        //每个病人的ICD确定疾病数为0
                        myDic.DiseaseEachPersonICDNum.Add(eachPerson, 0);
                        //每个病人的无ICD确定值疾病数为0
                        myDic.DiseaseEachPersonNotICDNum.Add(eachPerson, 0);
                        //跳过此人的循环
                        continue;
                    }
                    else
                    {
                        //此病人有诊断
                        foreach (var eachDisease in diseaseResult)
                        {
                            //不论有没有确定的ICD值，都要增加总的疾病数量
                            if (myDic.DiseaseEachPersonNum.ContainsKey(eachPerson))
                            {
                                myDic.DiseaseEachPersonNum[eachPerson]++;
                            }
                            else
                            {
                                //第一次统计此病人的疾病总量
                                myDic.DiseaseEachPersonNum.Add(eachPerson, 1);
                            }
                            //诊断有确定ICD值，相应的疾病ICD值加1
                            if (eachDisease.diagcode != null)
                            {
                                //区分患者的年龄
                                string sep_DiseaseAge = eachDisease.diagcode.ToString() + "，" + PatatienAgeRange ;
                                //区分患者的性别
                                string sep_DiseaseSex = eachDisease.diagcode.ToString() + "，" + patientSex ;
                                //区分患者的年龄和性别
                                string sep_DisesseSexAge = eachDisease.diagcode.ToString() + "，" + patientSex + "，" + PatatienAgeRange; ;
                                //统计全体疾病
                                if (myDic.Disease.ContainsKey(eachDisease.diagcode.ToString()))
                                {
                                    //已经遇到此种ICD疾病了
                                    myDic.Disease[eachDisease.diagcode]++;
                                }
                                else
                                {
                                    //第一次统计此种ICD值疾病的数量
                                    myDic.Disease.Add(eachDisease.diagcode, 1);
                                }
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
                                //分年龄和性别的
                                if (myDic.DiseaseAgeSex.ContainsKey(sep_DisesseSexAge))
                                {
                                    //此年龄和性别分组
                                    myDic.DiseaseAgeSex[sep_DisesseSexAge]++;
                                }
                                else
                                {
                                    myDic.DiseaseAgeSex.Add(sep_DisesseSexAge, 1);
                                }

                                //有确定的ICD值，那每个人的ICD确定诊断数目加1
                                if (myDic.DiseaseEachPersonICDNum.ContainsKey(eachPerson))
                                {
                                    myDic.DiseaseEachPersonICDNum[eachPerson]++;
                                }
                                else
                                {
                                    //第一次统计此人的ICD值
                                    myDic.DiseaseEachPersonICDNum.Add(eachPerson, 1);
                                }
                            }
                            //诊断没有确定的ICD值
                            else
                            {
                                //没有确定的ICD值，那此人的ICD不确定诊断数目加1
                                if (myDic.DiseaseEachPersonNotICDNum.ContainsKey(eachPerson))
                                {
                                    myDic.DiseaseEachPersonNotICDNum[eachPerson]++;
                                }
                                else
                                {
                                    //第一次统计此人的非ICD值
                                    myDic.DiseaseEachPersonNotICDNum.Add(eachPerson, 1);
                                }
                            }
                        }
                    }
                }//确定此病人有诊断的结束
                catch
                {
                    System.Windows.Forms.MessageBox.Show("遍历查询到的患者时出现错误");
                }
                
            }//Foreach查询到的所有患者循环
            return "成功执行统计";
        }//结束public string statistics(DateTime startDate, DateTime endDate)
    }//结束public class StatisticsService
}
