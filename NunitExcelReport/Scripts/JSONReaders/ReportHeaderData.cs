using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NunitExcelReport.Scripts.JSONReaders
{
    public class ReportHeaderData
    {


        public string Test_report_data { get; set; }
        public string Project_name_data { get; set; }
        public string Project_id_data { get; set; }
        public string Release_data { get; set; }

        public string Tester_name_data { get; set; }
        public string Test_report_notes_data { get; set; }
    }
}
