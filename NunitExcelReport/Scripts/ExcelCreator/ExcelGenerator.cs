using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using NunitExcelReport.Interfaces;
using NunitExcelReport.Scripts.ExcelCreator;
using NunitExcelReport.Scripts.JSONReaders;



namespace NunitExcelReport.Scripts.ExcelCreator;

public class ExcelGenerator
{
    private HeaderJsonReader jsonReader;
    private ExcelHeaderDataJsonReader ejsonHeaderDataReader;
    private IExcelHeaderCreator eHeaderCreator;

    private IExcelSummaryCreator excelSummaryCreator;
    private IExcelDetailedResultCreator excelDetailedResultCreator;
    private ExcelHeaderKeyJsonReader eHeaderKeyJsonReader;
    private MainConfigData mainConfigData;

    private static readonly NLog.Logger nlogger = NLog.LogManager.GetCurrentClassLogger();

    public ExcelGenerator()
    {
        var currentDirectory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
        string projectDirectory = currentDirectory.Parent.Parent.Parent.FullName;
        string configFile = projectDirectory + "\\Resources\\excel_report_generator_config.json";
        nlogger.Info("config file path: " + configFile);
        jsonReader = new HeaderJsonReader();

        MainConfigJsonReader mainConfigJsonReader = new MainConfigJsonReader();
        mainConfigData = mainConfigJsonReader.createJsonData(jsonReader.readJsonFile(configFile));
    }


    public void createReport(TestResultDataCollector result, TestContext context)
    {

        nlogger.Info("Reading excel headers keys.");
        nlogger.Info("Excel_header_Key_filePath: " + mainConfigData.Excel_header_key_filePath);
        nlogger.Info("Excel_header_data_filepath: " + mainConfigData.Excel_header_data_filepath);
        eHeaderKeyJsonReader = new ExcelHeaderKeyJsonReader();
        ExcelHeaderKeys excelHeaderkeys = eHeaderKeyJsonReader.createJsonData(jsonReader.readJsonFile(mainConfigData.Excel_header_key_filePath));

        ejsonHeaderDataReader = new ExcelHeaderDataJsonReader();
        ReportHeaderData reportHeaderData = ejsonHeaderDataReader.createJsonData(jsonReader.readJsonFile(mainConfigData.Excel_header_data_filepath));

        eHeaderCreator = new ExcelHeaderCreator();

        nlogger.Info("Reading excel eaders data.");
        excelSummaryCreator = new ExcelSummaryCreator();
        excelDetailedResultCreator = new ExcelDetailedResultCreator();


        var wbook = new XLWorkbook();
        var ws = wbook.Worksheets.Add(mainConfigData.Excel_report_worksheet_name);
    
        eHeaderCreator.createExcelHeader(ws, excelHeaderkeys, reportHeaderData);
        nlogger.Info("Create excel header.");



        excelSummaryCreator.createExcelSummary(ws, excelHeaderkeys, context);
        nlogger.Info("Create excel summary.");


        excelDetailedResultCreator.createDetailedresult(ws, result, excelHeaderkeys);


        nlogger.Info("Create excel deateled esult.");

        string excel_date_name = DateTime.Now.ToString("yyyy-MM-dd-h-mm-ss");
        // var lastColumnused=ws.LastColumnUsed();
        var columnUsed = ws.ColumnsUsed();
        columnUsed.AdjustToContents();
        string excel_name = mainConfigData.Excel_report_path + mainConfigData.Excel_report_name + "_" + excel_date_name + mainConfigData.Excel_report_extension;
        nlogger.Info("Save excel to: " + excel_name);

        wbook.SaveAs(excel_name);
    }






}