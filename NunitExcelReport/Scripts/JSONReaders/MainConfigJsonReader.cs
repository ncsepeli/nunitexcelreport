using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NunitExcelReport.Scripts.JSONReaders
{
    internal class MainConfigJsonReader: HeaderJsonReader
    {
        public MainConfigData createJsonData(string jsonString)
        {
            MainConfigData datamodel = new MainConfigData();

            try
            {
                datamodel = JsonConvert.DeserializeObject<MainConfigData>(jsonString);
            }
            catch (JsonReaderException e)
            {
                Console.WriteLine("Something wrong with the test data! " + e.Message);
            }
            return datamodel;
        }
    }
}

