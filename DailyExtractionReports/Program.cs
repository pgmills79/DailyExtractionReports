
using DailyExtractionReports;


Console.WriteLine("Running daily extraction files creation....");


//Retrieve Possible Duplicates
Console.WriteLine("Creating potential duplicates file....");
var possibleDuplicates = DatabaseHandler.GetPossibleDuplicates().ToList();
ExcelHandler.ExportDuplicatesToExcel(possibleDuplicates);

Console.WriteLine("Creating pending samples file....");
//Retrieve Pending Samples
var pendingSamples = DatabaseHandler.GetPendingSamples();
ExcelHandler.ExportPendingSamplesToExcel(pendingSamples);

Console.WriteLine("Daily extraction files have been created....");
Console.WriteLine("Press any key to close program");
Console.ReadKey();


	     
	     

