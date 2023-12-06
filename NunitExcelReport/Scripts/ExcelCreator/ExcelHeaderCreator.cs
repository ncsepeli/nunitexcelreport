using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Interfaces;
using NunitExcelReport.Scripts.JSONReaders;

namespace NunitExcelReport.Scripts.ExcelCreator
{
    public class ExcelHeaderCreator: IExcelHeaderCreator
    {

        public void createExcelHeader(IXLWorksheet ws, ExcelHeaderKeys cheaders, ReportHeaderData dataValue)
        {

            ws.Cell("A1").Value = cheaders.Test_report_key;
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("B1").Value = dataValue.Test_report_data;

            ws.Cell("A2").Value = cheaders.Project_name_key;
            ws.Cell("A2").Style.Font.Bold = true;
            ws.Cell("B2").Value = dataValue.Project_name_data;

            ws.Cell("A3").Value = cheaders.Project_id_key;
            ws.Cell("A3").Style.Font.Bold = true;
            ws.Cell("B3").Value = dataValue.Project_id_data;

            ws.Cell("A4").Value = cheaders.Release_key;
            ws.Cell("A4").Style.Font.Bold = true;
            ws.Cell("B4").Value = dataValue.Release_data;

            ws.Cell("A5").Value = cheaders.Tester_name_key;
            ws.Cell("A5").Style.Font.Bold = true;
            ws.Cell("B5").Value = dataValue.Tester_name_data;

            ws.Cell("A6").Value = cheaders.Test_report_notes_key;
            ws.Cell("A6").Style.Font.Bold = true;
            ws.Cell("B6").Value = dataValue.Test_report_notes_data;
        }
    }
}
