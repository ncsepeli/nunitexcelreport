using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Scripts.ExcelCreator;

namespace NunitExcelReport.Interfaces
{
    public interface IExcelGenerator
    {
        void createReport(TestResultDataCollector result, TestContext context);
    }
}
