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
    
    public partial class htaskDepImage
    {
        public string gb2260 { get; set; }
        public string checkcode { get; set; }
        public string taskcode { get; set; }
        public string deptcode { get; set; }
        public string seq { get; set; }
        public string ifdel { get; set; }
        public Nullable<System.DateTime> checkdate { get; set; }
        public string sampleno { get; set; }
        public string ifpath { get; set; }
        public string imagepath { get; set; }
        public byte[] imagedata { get; set; }
        public string asmdesc { get; set; }
        public string tag { get; set; }
    }
}