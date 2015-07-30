using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PatientDataExport.Package
{
    public class StatisticsParameters
    {
        public string workunit;
        public DateTime startDate;
        public DateTime endDate;
        public StatisticsParameters()
        {
            startDate = DateTime.Now.Date;
            endDate = DateTime.Now.Date.AddYears(1);
            workunit = "";
        }
    }
}
