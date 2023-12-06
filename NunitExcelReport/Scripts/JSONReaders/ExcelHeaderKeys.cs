using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Scripts;

namespace NunitExcelReport.Scripts.JSONReaders
{
    public class ExcelHeaderKeys
    {
        public string Test_report_key { get; set; }
        public string Project_name_key { get; set; }
        public string Project_id_key { get; set; }
        public string Release_key { get; set; }

        public string Tester_name_key { get; set; }
        public string Test_report_notes_key { get; set; }
        public string Test_all_result_key { get; set; }
        public string Test_passed_key { get; set; }

        public string Test_not_run_key { get; set; }
        public string Test_failed_key { get; set; }
        public string Test_id_key { get; set; }
        public string Test_name_key { get; set; }

        public string Test_description_key { get; set; }
        public string Test_result_key { get; set; }
        public string Test_run_date_key { get; set; }
        public string Test_note_key { get; set; }
        public string Ticket_id_key { get; set; }
    }
}

