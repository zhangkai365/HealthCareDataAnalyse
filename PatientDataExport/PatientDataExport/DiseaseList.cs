using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//
using Excel = Microsoft.Office.Interop.Excel;

namespace PatientDataExport
{
    public class DiseaseList
    {

        public void Initialize(string ICDFileName,out Dictionary<string,string> ICDList)
        {
            string defaultICDStoreFileName = @"C:\Users\Win7x64_20140606\Desktop\2014年查体工作总结（修改201408080753）\疾病总表\2014ICD2217.xlsx";
            if (ICDFileName == "") ICDFileName = defaultICDStoreFileName;
            try
            {
                Excel.Application ICDExcel = new Excel.Application();
                Excel.Workbook ICDWorkbook = ICDExcel.Workbooks.Open(ICDFileName);
                Excel.Worksheet ICDWorksheet = ICDWorkbook.Worksheets[1];
                ICDList = new Dictionary<string, string>();
                int count_empty = 0;
                string tempICDCode;
                string tempICDName;
                for (int i = 1; i < 200; i++)
                {
                    try
                    {
                        tempICDName = ICDWorksheet.get_Range("A" + i).Text;
                        tempICDCode = ICDWorksheet.get_Range("B" + i).Text;
                        if (!ICDList.Keys.Contains(tempICDCode))
                        {
                            ICDList.Add(tempICDCode,tempICDName);
                        }
                    }
                    catch
                    {
                        count_empty++;
                        continue;
                    }
                    if (count_empty > 2) break;
                }
                ICDWorkbook.Close();
                ICDExcel.Quit();
            }
            catch
            {
                ICDList = new Dictionary<string, string>();
                ICDList.Add("读取疾病列表出错", "Error");
            }
        }
    }
}
