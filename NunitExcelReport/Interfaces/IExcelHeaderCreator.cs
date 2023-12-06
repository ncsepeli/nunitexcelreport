using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Scripts.JSONReaders;

namespace NunitExcelReport.Interfaces
{
    internal interface IExcelHeaderCreator
    {
        void createExcelHeader(IXLWorksheet ws, ExcelHeaderKeys cheaders, ReportHeaderData dataValue);
    }
}
