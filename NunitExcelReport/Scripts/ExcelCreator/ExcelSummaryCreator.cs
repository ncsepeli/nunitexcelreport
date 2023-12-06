using ClosedXML.Excel;
using DocumentFormat.OpenXml.InkML;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Interfaces;
using NunitExcelReport.Scripts.JSONReaders;
using NUnit.Framework;

namespace NunitExcelReport.Scripts.ExcelCreator
{
    public class ExcelSummaryCreator: IExcelSummaryCreator
    {
        private static readonly NLog.Logger nlogger = NLog.LogManager.GetCurrentClassLogger();
        public void createExcelSummary(IXLWorksheet ws, ExcelHeaderKeys headerKeys, TestContext context)
        {
            int rowNum = ws.LastRowUsed().RowNumber() + 2;
            ws.Cell("A"+ rowNum).Value = headerKeys.Test_all_result_key;
            ws.Cell("A"+ rowNum).Style.Font.Bold = true;
            ws.Cell("B"+ rowNum).Value = context.Result.FailCount + context.Result.PassCount + context.Result.SkipCount;
            nlogger.Info("context.Result sum: " + (context.Result.FailCount + context.Result.PassCount + context.Result.SkipCount));
            ws.Cell("A"+ (rowNum+1)).Value = headerKeys.Test_passed_key;
            ws.Cell("A"+(rowNum + 1)).Style.Font.Bold = true;
            ws.Cell("B" + (rowNum + 1)).Value = context.Result.PassCount;

            nlogger.Info("context.Result.PassCount: " + context.Result.PassCount);


            ws.Cell("A" + (rowNum + 2)).Value = headerKeys.Test_failed_key;
            ws.Cell("A" + (rowNum + 2)).Style.Font.Bold = true;
            ws.Cell("B"+(rowNum + 2)).Value = context.Result.FailCount;
            nlogger.Info("context.Result.FailCount: " + context.Result.FailCount);

            ws.Cell("A"+(rowNum + 3)).Value = headerKeys.Test_not_run_key;
            ws.Cell("A" + (rowNum + 3)).Style.Font.Bold = true;
            ws.Cell("B" + (rowNum + 3)).Value = context.Result.SkipCount;
            nlogger.Info("context.Result.SkipCount: " + context.Result.SkipCount);

        }
    }
}
