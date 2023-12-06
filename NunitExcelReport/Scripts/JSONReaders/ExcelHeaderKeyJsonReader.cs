using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NunitExcelReport.Scripts.JSONReaders
{
    public class ExcelHeaderKeyJsonReader : HeaderJsonReader
    {
        public ExcelHeaderKeys createJsonData(string jsonString)
        {
            ExcelHeaderKeys datamodel = new ExcelHeaderKeys();

            try
            {
                datamodel = JsonConvert.DeserializeObject<ExcelHeaderKeys>(jsonString);
            }
            catch (JsonReaderException e)
            {
                Console.WriteLine("Something wrong with the test data! " + e.Message);
            }
            return datamodel;
        }
    }
}
