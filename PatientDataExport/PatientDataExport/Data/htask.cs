//------------------------------------------------------------------------------
// <auto-generated>
//     此代码已从模板生成。
//
//     手动更改此文件可能导致应用程序出现意外的行为。
//     如果重新生成代码，将覆盖对此文件的手动更改。
// </auto-generated>
//------------------------------------------------------------------------------

namespace PatientDataExport.Data
{
    using System;
    using System.Collections.Generic;
    
    public partial class htask
    {
        public string taskcode { get; set; }
        public string gb2260 { get; set; }
        public string membtype { get; set; }
        public string taskname { get; set; }
        public string taskyear { get; set; }
        public string taskrunid { get; set; }
        public string password { get; set; }
        public string packpass { get; set; }
        public string hiscode { get; set; }
        public Nullable<int> taskgrps { get; set; }
        public Nullable<int> taskmembs { get; set; }
        public Nullable<int> summale { get; set; }
        public Nullable<int> sumfemale { get; set; }
        public Nullable<int> finishgrps { get; set; }
        public Nullable<int> finishmembs { get; set; }
        public Nullable<int> finishrate { get; set; }
        public Nullable<int> leftdeptsum { get; set; }
        public Nullable<int> newsignsum { get; set; }
        public Nullable<System.DateTime> builddate { get; set; }
        public string builder { get; set; }
        public Nullable<System.DateTime> startdate { get; set; }
        public Nullable<System.DateTime> stopdate { get; set; }
        public Nullable<decimal> totalfee { get; set; }
        public Nullable<decimal> addfee { get; set; }
        public Nullable<decimal> orderfee { get; set; }
        public string payid { get; set; }
        public string ifdel { get; set; }
        public string ifhide { get; set; }
        public string ifcipher { get; set; }
        public string ifbase { get; set; }
        public string remark { get; set; }
        public string tag { get; set; }
        public string tasktype { get; set; }
        public string ifsettle { get; set; }
        public string autoconfirm { get; set; }
        public string todo_unsettle { get; set; }
        public string autoxfee { get; set; }
    }
}