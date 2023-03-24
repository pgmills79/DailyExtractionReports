
using DailyExtractionReports;

	

Console.WriteLine("Running daily extraction files creation....");

var fileDate = ExcelHandler.GetDateToAppendToFileName();

//Retrieve Possible Duplicates
Console.WriteLine("Creating potential duplicates file....");
var possibleDuplicates = DatabaseHandler.GetPossibleDuplicates().ToList();
var fileName = $"{ExcelHandler.BaseDirectory}Daily_Duplicates_{fileDate}.xlsx";
ExcelHandler.ExportDuplicatesToExcel(possibleDuplicates, fileName);

Console.WriteLine("Creating pending samples file....");
//Retrieve Pending Samples
var pendingSamples = DatabaseHandler.GetPendingSamples();
fileName = $"{ExcelHandler.BaseDirectory}clinmicro_pending_list_{fileDate}.xlsx";
ExcelHandler.ExportPendingSamplesToExcel(pendingSamples, fileName);

Console.WriteLine("Daily extraction files have been created....");
Console.WriteLine("Press any key to close program");
Console.ReadKey();


	     
	     

