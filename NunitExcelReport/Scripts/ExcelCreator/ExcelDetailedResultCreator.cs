using ClosedXML.Excel;
using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Scripts.JSONReaders;
using NunitExcelReport.Interfaces;

namespace NunitExcelReport.Scripts.ExcelCreator
{
    public class ExcelDetailedResultCreator: IExcelDetailedResultCreator
    {
        private static readonly NLog.Logger nlogger = NLog.LogManager.GetCurrentClassLogger();
        public void createDetailedresult(IXLWorksheet ws, TestResultDataCollector result, ExcelHeaderKeys headerKeys)
        {
            nlogger.Info("createDetailedresult");
            int cellRow =  ws.LastRowUsed().RowNumber() + 3, cell_row_number = 0, cellRowHeader=cellRow-1;
            List<ResultData> resultList = result.GetTestResultList();
            TestContext testContext;

            nlogger.Info("write_part_result_into_excel2");
            ws.Cell("A" + cellRowHeader).Value = headerKeys.Test_name_key;
            ws.Cell("A" + cellRowHeader).Style.Font.Bold = true;
            ws.Cell("B" + cellRowHeader).Value = headerKeys.Test_description_key;
            ws.Cell("B" + cellRowHeader).Style.Font.Bold = true;
            ws.Cell("C" + cellRowHeader).Value = headerKeys.Test_id_key;
            ws.Cell("C" + cellRowHeader).Style.Font.Bold = true;
            ws.Cell("D" + cellRowHeader).Value = headerKeys.Test_result_key;
            ws.Cell("D" + cellRowHeader).Style.Font.Bold = true;
            ws.Cell("E" + cellRowHeader).Value = headerKeys.Test_run_date_key;
            ws.Cell("E" + cellRowHeader).Style.Font.Bold = true;
            ws.Cell("F" + cellRowHeader).Value = headerKeys.Test_note_key;
            ws.Cell("F" + cellRowHeader).Style.Font.Bold = true;
            ws.Cell("G" + cellRowHeader).Value = headerKeys.Ticket_id_key;
            ws.Cell("G" + cellRowHeader).Style.Font.Bold = true;





            for (int j = 0; j < resultList.Count; j++)
            {
                // nlogger.Info("i: " + i);
                nlogger.Info("j: " + j);
                testContext = resultList[j].resultContext;
                cell_row_number = cellRow + j;
                ws.Cell("A" + cell_row_number).Value = testContext.Test.FullName;
                nlogger.Info("A: " + cell_row_number + " A" + cell_row_number + " " + resultList[j].resultContext.Test.FullName);
                ws.Cell("B" + cell_row_number).Value = (XLCellValue)checkResultData((string)testContext.Test.Properties.Get("Description"));  //datamodel.Test_description_key;
                nlogger.Info("B: " + cell_row_number + " B" + cell_row_number + " " + (XLCellValue)checkResultData((string)testContext.Test.Properties.Get("Description")));
                ws.Cell("C" + cell_row_number).Value = checkResultData(testContext.Test.ID);  // datamodel.Test_id_key;
                nlogger.Info("C: " + cell_row_number + " C" + cell_row_number + " " + checkResultData(testContext.Test.ID));
                ws.Cell("D" + cell_row_number).Value = checkResultData(Enum.GetName(typeof(TestStatus), testContext.Result.Outcome.Status));//  datamodel.Test_result_key;
                nlogger.Info("D: " + cell_row_number + " D" + cell_row_number + " " + checkResultData(Enum.GetName(typeof(TestStatus), testContext.Result.Outcome.Status)));
                ws.Cell("E" + cell_row_number).Value = checkResultData(resultList[j].resultDate.ToString("yyyy-MM-dd-h-mm-ss")); //   datamodel.Test_run_date_key;
                nlogger.Info("E: " + cell_row_number + " E" + cell_row_number + " " + checkResultData(resultList[j].resultDate.ToString("yyyy-MM-dd-h-mm-ss")));
                ws.Cell("F" + cell_row_number).Value = (XLCellValue)checkResultData((string)testContext.Test.Properties.Get("Note"));
                nlogger.Info("F: " + cell_row_number + " F" + cell_row_number + " " + (XLCellValue)checkResultData((string)testContext.Test.Properties.Get("Note")));
                ws.Cell("G" + cell_row_number).Value = (XLCellValue)checkResultData((string)testContext.Test.Properties.Get("ID"));
                nlogger.Info("G: " + cell_row_number + " G" + cell_row_number + " " + (XLCellValue)checkResultData((string)testContext.Test.Properties.Get("ID")));
            }

            nlogger.Info("write_part_result_into_excel3");
        }

        string checkResultData(string result)
        {

            if (result == null) return "no_data";
            else return result;
        }
    }
}
