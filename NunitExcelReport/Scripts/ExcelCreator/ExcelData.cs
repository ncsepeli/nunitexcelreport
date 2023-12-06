using DocumentFormat.OpenXml.Office2010.Excel;
using NUnit.Framework.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NunitExcelReport.Scripts;
using NUnit.Framework;

namespace NunitExcelReport.Scripts.ExcelCreator

{
    public class TestResultDataCollector
    {
        private List<ResultData> testContextList;


        public TestResultDataCollector()
        {
            if (testContextList == null)
            {
                testContextList = new List<ResultData>();
            }
        }
        public void AddToReport(TestContext context)
        {

            ResultData resultData = new ResultData();
            resultData.resultContext = context;
            resultData.resultDate = DateTime.Now;
            testContextList.Add(resultData);
        }

        public List<ResultData> GetTestResultList()
        { return testContextList; }

    }
}
