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
            if (ICDFileName == "")
            {
                System.Windows.Forms.MessageBox.Show("没有指定疾病总表的路径，程序错误！");
                throw new Exception();
            } 
            try
            {
                Excel.Application ICDExcel = new Excel.Application();
                Excel.Workbook ICDWorkbook = ICDExcel.Workbooks.Open(ICDFileName);
                Excel.Worksheet ICDWorksheet = ICDWorkbook.Worksheets[1];
                ICDList = new Dictionary<string, string>();
                int count_empty = 0;
                string tempICDCode;
                string tempICDName;
                for (int i = 1; i < 2000; i++)
                {
                    try
                    {
                        tempICDCode = ICDWorksheet.get_Range("A" + i).Text;
                        tempICDName = ICDWorksheet.get_Range("B" + i).Text;
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
