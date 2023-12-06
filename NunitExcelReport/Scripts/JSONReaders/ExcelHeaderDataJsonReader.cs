using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NunitExcelReport.Scripts.JSONReaders
{
    public class ExcelHeaderDataJsonReader : HeaderJsonReader
    {
        public ReportHeaderData createJsonData(string jsonString)
        {
            ReportHeaderData datamodel = new ReportHeaderData();

            try
            {
                datamodel = JsonConvert.DeserializeObject<ReportHeaderData>(jsonString);
            }
            catch (JsonReaderException e)
            {
                Console.WriteLine("Something wrong with the test data! " + e.Message);
            }
            return datamodel;
        }
    }
    
   
}
