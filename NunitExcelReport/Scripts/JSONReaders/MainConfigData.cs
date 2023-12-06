using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NunitExcelReport.Scripts.JSONReaders
{
    public class MainConfigData
    {
        /**With the excel_report_generator_config.json can configure the excel report.
         * The excel generator searches for the main config json file in the Solution\Resources directory.**/

        public string Excel_report_path { get; set; }// The  Excel_report_path tells where to put the generated excel file.
        public string Excel_report_name { get; set; }// The Excel_report_name defines the report name. It will get timestamp into its name.
        public string Excel_report_extension { get; set; } //Define the excel extension. If you give wrong extension name, the excel will not created.
        public string Excel_header_key_filePath { get; set; }// This json defines the data types of the report: e.g: Tester_name, Project_name, bold in the example excel.
        public string Excel_header_data_filepath { get; set; }//This json defines the main data of the report. e.g: Tester Name: John Winchester
        public string Excel_report_worksheet_name { get; set; }//This will be the excel worksheet name.Only one excel sheet created during the test.
      
    }
}

