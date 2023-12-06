using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Scripts.ExcelCreator;
using NunitExcelReport.Scripts.JSONReaders;

namespace NunitExcelReport.Interfaces
{
    public interface IExcelDetailedResultCreator
    {
         void createDetailedresult(IXLWorksheet ws, TestResultDataCollector result, ExcelHeaderKeys headerKeys);
    }
}
