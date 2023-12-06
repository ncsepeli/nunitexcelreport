using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NunitExcelReport.Scripts.JSONReaders
{
    public class HeaderJsonReader
    {
        private static readonly NLog.Logger nlogger = NLog.LogManager.GetCurrentClassLogger();
        public string readJsonFile(string filepath)
        {

            string jsonString = "";

            try
            {
                StreamReader r = new StreamReader(filepath);
                jsonString = r.ReadToEnd();

            }
            catch (DriveNotFoundException e)
            {
                nlogger.Error(e.Message);
                nlogger.Error("Cannot read the file! The filepath was: " + filepath);
            }

            return jsonString;
        }
        public string createJsonData(string jsonString)
        {

            return "no_data";
        }
    }
}

