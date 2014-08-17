﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//
using Excel = Microsoft.Office.Interop.Excel;

namespace PatientDataExport
{
    public class OutputExcelService
    {
        /// <summary>
        /// 确定疾病的名称
        /// </summary>
        /// <param name="allICDDisease">全部疾病的列表</param>
        /// <param name="ICDCode">待确定的疾病编码</param>
        /// <returns></returns>
        public string DiagnosisName(Dictionary<string, string> allICDDisease, string ICDCode)
        {
            string[] ICDcodeSplite = null;
            ICDcodeSplite =  ICDCode.Split('，');
            if (ICDcodeSplite == null) return "分割后字符为空";
            string tempICD = ICDcodeSplite[0];
            if (tempICD == null) return "字符串数组为空";
            try
            {
                var ICD = (from s1 in allICDDisease where s1.Key.Contains(tempICD) select s1).First();
                return ICD.Value;
            }
            catch
            {
                return "无此诊断编码";
            }
        }

        public string OutputExcel(string FilePath, Dic myDic, Dictionary<string, string> List_Disease)
        {
            //设置ICD存储文件的路径及文件名
            //string ICDFileNameStore = @"C:\Users\Win7x64_20140606\Desktop\2014年查体工作总结（修改201408080753）\疾病总表\ICD2.xls";

            Excel.Application myExcel = new Excel.Application();
            myExcel.Visible = false;
            //存储统计结果
            Excel.Workbook myWorkbook = myExcel.Workbooks.Add(true);
            Excel.Worksheet myWorkSheet = myWorkbook.Worksheets[1];

            //存储主要的疾病的ICD诊断号码的Excel
            //Excel.Workbook ICDWorkbook = myExcel.Workbooks.Add(true);
            //Excel.Worksheet ICDWorksheet = ICDWorkbook.Worksheets[1];

            //输出统计结果
            //输出处理的人数
            myWorkSheet.Cells[1, 3] = "统计的人数";
            myWorkSheet.Cells[1, 4] = myDic.NumAll;
            //男性人数人数
            myWorkSheet.Cells[1, 5] = "男";
            myWorkSheet.Cells[1, 6] = myDic.NumSex["男"];
            //女性人数
            myWorkSheet.Cells[1, 7] = "女";
            myWorkSheet.Cells[1, 8] = myDic.NumSex["女"];
            //不同年龄段的人数范围
            int tempCountRangeNum = 0;
            foreach (var eachRangeNum in myDic.NumAge)
            {
                myWorkSheet.Cells[2, tempCountRangeNum + 1] = eachRangeNum.Key.ToString();
                myWorkSheet.Cells[2, tempCountRangeNum + 2] = eachRangeNum.Value.ToString();
                tempCountRangeNum = tempCountRangeNum + 2;
            }
            //不同年龄段男女数目
            int tempCountSexAgeNum = 0;
            foreach (var eachSexAgeNum in myDic.NumSexAge)
            {
                myWorkSheet.Cells[3, tempCountSexAgeNum + 1] = eachSexAgeNum.Key.ToString();
                myWorkSheet.Cells[3, tempCountSexAgeNum + 2] = eachSexAgeNum.Value.ToString();
                tempCountSexAgeNum = tempCountSexAgeNum + 2;
            }

            //总疾病的数量
            myWorkSheet.Cells[5, 1] = "总疾病的数量";
            int totalDiseaseNum = 0;
            foreach (var total_eachPersonDiseaseNum in myDic.DiseaseEachPersonNum)
            {
                totalDiseaseNum = totalDiseaseNum + total_eachPersonDiseaseNum.Value;
            }
            myWorkSheet.Cells[5, 2] = totalDiseaseNum;

            //总ICD诊断数量
            myWorkSheet.Cells[5, 4] = "总ICD诊断的数量";
            int totalICDDiseaseNum = 0;
            foreach (var total_eachPersonICDDiseaseNum in myDic.DiseaseEachPersonICDNum)
            {
                totalICDDiseaseNum = totalICDDiseaseNum + total_eachPersonICDDiseaseNum.Value;
            }
            myWorkSheet.Cells[5, 5] = totalICDDiseaseNum;

            //总非ICD诊断数量
            myWorkSheet.Cells[5, 6] = "总非ICD诊断的数量";
            int totalNotICDDiseaseNum = 0;
            foreach (var total_eachPersonNotICDDiseaseNum in myDic.DiseaseEachPersonNotICDNum)
            {
                totalNotICDDiseaseNum = totalNotICDDiseaseNum + total_eachPersonNotICDDiseaseNum.Value;
            }
            myWorkSheet.Cells[5, 7] = totalNotICDDiseaseNum;

            //各个ICD诊断的数量
            //简化ICD表头
            //性别字符串集合
            string[] GroupSex = { "男", "女" };
            //疾病诊断
            string[] GroupTitle = { "ICD诊断名称", "ICD诊断编码", "此ICD诊断发病数量" };
            int titleRangeAll = 1;
            foreach (string preTitleSex in GroupSex)
            {
                foreach (string titleDisease in GroupTitle)
                {
                    myWorkSheet.Cells[6, titleRangeAll++] = preTitleSex + "，" + titleDisease;
                }
            }
            //筛选此范围内前50位的诊断，分男女
            //控制横向移动的变量
            int countX_Top20all = 0;
            foreach (string top20allSex in GroupSex)
            {
                //控制纵向移动的变量
                int countY_Top20all = 0;
                try
                {
                    var top20male = (from temptop20all in myDic.DiseaseSex
                                     where temptop20all.Key.Contains(top20allSex)
                                     orderby temptop20all.Value
                                     descending
                                     select temptop20all).Take(100);
                    if (top20male == null)
                    {
                        myWorkSheet.Cells[7, 2 + countX_Top20all] = "空白";
                        myWorkSheet.Cells[7, 3 + countX_Top20all] = "空白";
                        //ICDWorksheet.Cells[1, 2 + countX_Top20all] = "空白";
                        //ICDWorksheet.Cells[1, 3 + countX_Top20all] = "空白";
                    }
                    else
                    {
                        foreach (var eachICDDiseaseNum in top20male)
                        {
                            //做出统计
                            myWorkSheet.Cells[7 + countY_Top20all, 1 + countX_Top20all] = DiagnosisName(List_Disease, eachICDDiseaseNum.Key.ToString());
                            myWorkSheet.Cells[7 + countY_Top20all, 2 + countX_Top20all] = eachICDDiseaseNum.Key.ToString();
                            myWorkSheet.Cells[7 + countY_Top20all, 3 + countX_Top20all] = eachICDDiseaseNum.Value.ToString();
                            //写入所有的诊断列表
                            //ICDWorksheet.Cells[1 + countY_Top20all, 2 + countX_Top20all] = eachICDDiseaseNum.Key.ToString();
                            //ICDWorksheet.Cells[1 + countY_Top20all, 3 + countX_Top20all] = eachICDDiseaseNum.Value.ToString();
                            //纵向移动
                            countY_Top20all++;
                        }
                    }
                }
                catch
                {
                    myWorkSheet.Cells[7, 2] = "异常";
                    myWorkSheet.Cells[7, 3] = "异常";
                    //ICDWorksheet.Cells[1, 2] = "异常";
                    //ICDWorksheet.Cells[1, 3] = "异常";
                }
                //横向移动
                countX_Top20all = countX_Top20all + 3;
            }

            //分年龄、性别ICD诊断的数量的标题，简化
            string[] GroupAge = { "<45", "45-50", "50-55","55-60", ">60" };
            //控制表头横向移动的变量
            int countX_titleSexAgeDisease = 0;
            foreach (string eachSex in GroupSex)
            {
                foreach (string eachAge in GroupAge)
                {
                    foreach (string title_Disease in GroupTitle)
                    {
                        myWorkSheet.Cells[6, 7 + countX_titleSexAgeDisease++] = eachSex + "，" + eachAge + "，" + title_Disease;
                    }
                }
            }

            //控制疾病列表的横向移动
            int countX2_SexAgeDisease = 0;

            //循环性别
            foreach (string eachSex in GroupSex)
            {
                //循环年龄
                foreach (string eachAge in GroupAge)
                {
                    //控制疾病移动的纵向移动,每次循环之内的疾病都清零
                    int countY2_SexAgeDisease = 0;
                    try
                    {
                        var top20DiseaseOfEachAge = (from tempTop20Disease in myDic.DiseaseAgeSex
                                                     where tempTop20Disease.Key.Contains(eachSex + "，" + eachAge)
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
            myExcel.Quit();
            return "成功输出结果";

        }//结束public string OutputExcel(string FilePath, Dic myDic)
    }//结束public class OutputExcelService
}