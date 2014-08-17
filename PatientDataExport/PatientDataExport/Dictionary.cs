using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PatientDataExport
{
    public class Dic
    {
        //总人数
        public int NumAll = 0;
        //统计不同年龄段的人数
        public Dictionary<string, int> NumAge = new Dictionary<string, int>();
        //统计不同性别的人数
        public Dictionary<string, int> NumSex = new Dictionary<string, int>();
        //统计不同年龄段和不同性别的人数
        public Dictionary<string, int> NumSexAge = new Dictionary<string, int>();

        //疾病统计
        public Dictionary<string, int> Disease = new Dictionary<string, int>();
        //分性别统计
        public Dictionary<string, int> DiseaseSex = new Dictionary<string, int>();
        //分年龄统计
        public Dictionary<string, int> DiseaseAge = new Dictionary<string, int>();
        //分年龄、性别统计
        public Dictionary<string, int> DiseaseAgeSex = new Dictionary<string, int>();

        //每个人的诊断数目统计
        public Dictionary<string, int> DiseaseEachPersonNum = new Dictionary<string, int>();
        //每个人的有确定ICD10疾病数目
        public Dictionary<string, int> DiseaseEachPersonICDNum = new Dictionary<string, int>();
        //每个人无确定ICD诊断数目
        public Dictionary<string, int> DiseaseEachPersonNotICDNum = new Dictionary<string, int>();
    }
}
