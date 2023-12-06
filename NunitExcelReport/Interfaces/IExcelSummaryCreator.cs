using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Scripts.JSONReaders;
using NUnit.Framework;

namespace NunitExcelReport.Interfaces
{
     public interface IExcelSummaryCreator
    {
        void createExcelSummary(IXLWorksheet ws, ExcelHeaderKeys headerKeys, TestContext context);
    }
}
