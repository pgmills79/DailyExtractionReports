using System.Diagnostics.CodeAnalysis;
using ClosedXML.Excel;
using DailyExtractionReports.Models;

namespace DailyExtractionReports;

public static class ExcelHandler
{
    public const string DuplicatesWorksheetName = "Duplicate Results";
    private const string PendingWorksheetName = "Pending Samples";
    public const string BaseDirectory = "C:\\Daily Query Report Files\\";

    public static void ExportDuplicatesToExcel(List<Duplicates> possibleDuplicates)
    {
        
        var fileDate = GetDateToAppendToFileName();
        var fileName = $"C:\\Daily Query Report Files\\Daily_Duplicates_{fileDate}.xlsx";

        var currentRow = CreateWorksheet(DuplicatesWorksheetName, out var workbook, out var worksheet);
        FormatWorksheet<Duplicates>(worksheet, currentRow);
        
        AddDuplicatesToWorksheet(possibleDuplicates, currentRow, worksheet);

        SaveContentToFile(worksheet, workbook, fileName);
    }

    public static void AddDuplicatesToWorksheet(IReadOnlyCollection<Duplicates> possibleDuplicates, int currentRow, IXLWorksheet? worksheet)
    {
        foreach (var possibleDuplicate in possibleDuplicates.Where(x => long.TryParse(x.SpecId, out _)))
        {
            //only add if its a true duplicate (It has multiple rows with same specimen id AND the same status)
            if (IsNotTrueDuplicate(possibleDuplicates, possibleDuplicate)) continue;

            //TODO: check if spec id has been listed in a previous extraction report (then we dont have to see if it has been emailed out)


            currentRow++;
            AddWorksheetValues(worksheet, currentRow, possibleDuplicate);
        }
    }

    public static string GetDateToAppendToFileName()
    {
        return DateTime.Now.Month.ToString().Length == 1 ? "0" + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Year
            : DateTime.Now.Month.ToString() + DateTime.Now.Day + DateTime.Now.Year;
    }

    public static int CreateWorksheet(string worksheetName, out XLWorkbook workbook, out IXLWorksheet? worksheet)
    {
        const int currentRow = 1;

        workbook = new XLWorkbook();

        worksheet = workbook.Worksheets.Add(worksheetName);
        return currentRow;
    }

    private static bool IsNotTrueDuplicate(IReadOnlyCollection<Duplicates> possibleDuplicates, Duplicates possibleDuplicate)
    {
        return possibleDuplicates.Count(x => x.SpecId.Equals(possibleDuplicate.SpecId)) <= 1
               || possibleDuplicates.Where(x => x.SpecId.Equals(possibleDuplicate.SpecId))
                   .Select(x => x.ProcessingStatus).Distinct().Count() != 1;
    }
    
    public static void ExportPendingSamplesToExcel(IEnumerable<PendingSamples> pendingSamples)
    {
        var fileDate = DateTime.Now.Month.ToString().Length == 1 ? "0" + DateTime.Now.Month + DateTime.Now.Day + DateTime.Now.Year
            : DateTime.Now.Month.ToString() + DateTime.Now.Day + DateTime.Now.Year;
        
        var fileName = $"{BaseDirectory}clinmicro_pending_list_{fileDate}.xlsx";

        var currentRow = CreateWorksheet(PendingWorksheetName, out var workbook, out var worksheet);
        FormatWorksheet<PendingSamples>(worksheet, currentRow);

        foreach (var pendingSample in pendingSamples)
        {
            currentRow++;
            AddWorksheetValues(worksheet, currentRow, pendingSample);
        }
        
        //save the worksheet
        SaveContentToFile(worksheet, workbook, fileName);
    }
    
    private static void SaveContentToFile(IXLWorksheet? worksheet, IXLWorkbook workbook, string fileName)
    {
        worksheet?.Columns().AdjustToContents();
        worksheet?.SheetView.FreezeRows(0);
        workbook.SaveAs(fileName);
    }
    
    public static void AddWorksheetValues<T>(IXLWorksheet? worksheet, int currentRow, T recordToAdd) 
        where T : class
    {
        var numberProperties = typeof(T).GetProperties().Length;
        var t = recordToAdd.GetType();
       
        for(var i = 0;i <= numberProperties;i++)
        {
            var prop = t.GetProperties().ElementAtOrDefault(i);
            if (worksheet != null) worksheet.Cell(currentRow, i + 1).Value = prop?.GetValue(recordToAdd)?.ToString();
        }
    }

    public static void FormatWorksheet<T>(IXLWorksheet? worksheet, int currentRow)
    {
        
        var properties = typeof(T).GetProperties();

        for(var i = 0;i <= properties.Length;i++)
        {
            if (worksheet == null) continue;
            worksheet.Cell(currentRow, i + 1).Value = properties.ElementAtOrDefault(i)?.Name;
            worksheet.Cell(currentRow, i + 1).Style.Font.Bold = true;
        }
    }

    public static IEnumerable<string> GetYesterdayDuplicateSpecimenIds()
    {
        var files = GetLastDaysDuplicateFileInfo();
        var duplicateWorksheet = GetDuplicateWorksheet(files);

        var duplicateSpecimenIds = new List<string>(GetDistinctSpecimenIdsFromWorksheet(duplicateWorksheet) ?? Array.Empty<string>());

        return duplicateSpecimenIds;
    }

    public static IEnumerable<string>? GetDistinctSpecimenIdsFromWorksheet(IXLWorksheet? duplicateWorksheet)
    {
        return duplicateWorksheet?.Range("A2:A100").CellsUsed().Select(c => c.Value.ToString()).Distinct().ToList();
    }

    public static IXLWorksheet? GetDuplicateWorksheet(FileSystemInfo? files)
    {
        var workbook = new XLWorkbook(files?.FullName);
        var duplicateWorksheet = workbook.Worksheet(1);
        return duplicateWorksheet;
    }

    public static FileInfo? GetLastDaysDuplicateFileInfo()
    {
        var d = new DirectoryInfo(BaseDirectory); //Assuming Test is your Folder

        var fileInfo = d.GetFiles("*Duplicates*").OrderByDescending(x => x.LastWriteTime)
            .FirstOrDefault(x => (DateTime.Today - x.LastWriteTime).TotalHours >= 12);

        return fileInfo;
    }
}