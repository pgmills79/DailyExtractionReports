using System.Collections;
using DailyExtractionReports;
using DailyExtractionReports.Models;
using FluentAssertions;

namespace DailyExtractionReportTests;

public class ExcelHandlerTests
{
    
    private const string EmptyValue = "Empty Value";
    
    [Fact]
    public void GetDateToAppendToFileName_Should_Return_String_With_Today_Date()
    {
        //Arrange
        var todaysDate = DateTime.Now;

        //Act
        var result = ExcelHandler.GetDateToAppendToFileName();
        
        //Assert
        result.Should().NotBeNullOrEmpty();
        result.Should().Contain(todaysDate.Month.ToString());
        result.Should().Contain(todaysDate.Day.ToString());
        result.Should().Contain(todaysDate.Year.ToString());

        if (Convert.ToInt16(todaysDate.Month) < 10)
            result.Should().StartWith("0");

    }
    
    [Fact]
    public void CreateWorksheet_Should_Create_Worksheet()
    {
        //Arrange
        const string worksheetName = "Test Worksheet";

        //Act
        var result = ExcelHandler.CreateWorksheet(worksheetName, out var workbook, out var worksheet);
        
        //Assert
        result.Should().Be(1);
        workbook.Should().NotBeNull();
        worksheet.Should().NotBeNull();
        workbook.Worksheets.Should().HaveCount(1);

    }
    
    [Fact]
    public void FormatWorksheet_Duplicates_Should_Format_With_Correct_Header_Columns()
    {
        //Arrange
        const int currentRow = 1;
        const string worksheetName = "Test Worksheet";
        var duplicatesType = GetDuplicateClassPropertyCount(out var propertyCount);
        ExcelHandler.CreateWorksheet(worksheetName, out _, out var worksheet);
        
        //Act
        ExcelHandler.FormatWorksheet<Duplicates>(worksheet, currentRow);
        
        //Assert
        worksheet?.CellsUsed().Should().NotBeNullOrEmpty();
        worksheet?.CellsUsed().Should().HaveCount(propertyCount);
        worksheet?.CellsUsed().Select(c => c.Value.ToString()).Distinct().ToList()
            .Should().BeEquivalentTo(duplicatesType.Properties().Select(x => x.Name).ToList());
        

    }

    private static Type GetDuplicateClassPropertyCount(out int propertyCount)
    {
        var duplicatesType = typeof(Duplicates);
        propertyCount = duplicatesType.Properties().Count();
        return duplicatesType;
    }

    [Fact]
    public void FormatWorksheet_PendingSamples_Should_Format_With_Correct_Header_Columns()
    {
        //Arrange
        const int currentRow = 1;
        const string worksheetName = "Test Worksheet";
        var pendingSamplesType = typeof(PendingSamples);
        var propertyCount = pendingSamplesType.Properties().Count();
        ExcelHandler.CreateWorksheet(worksheetName, out _, out var worksheet);
        
        //Act
        ExcelHandler.FormatWorksheet<PendingSamples>(worksheet, currentRow);
        
        //Assert
        worksheet?.CellsUsed().Should().NotBeNullOrEmpty();
        worksheet?.CellsUsed().Should().HaveCount(propertyCount);
        worksheet?.CellsUsed().Select(c => c.Value.ToString()).Distinct().ToList()
            .Should().BeEquivalentTo(pendingSamplesType.Properties().Select(x => x.Name).ToList());
        

    }

    [Fact]
    public void GetLastDaysDuplicateFileInfo_Should_Return_Yesterdays_File()
    {
        //Arrange

        //Act
        var info = ExcelHandler.GetLastDaysDuplicateFileInfo();
        
        //Assert
        info?.Should().NotBeNull();
        info?.Name.Should().Contain(".xlsx");
        info?.LastWriteTime.Should().BeBefore(DateTime.Today);
    }
    
    [Fact]
    public void GetYesterdayDuplicateSpecimenIds_Should_Return_Yesterday_Duplicate_SpecimenIds()
    {
        //Arrange
        var yesterdayFile = ExcelHandler.GetLastDaysDuplicateFileInfo();
        var duplicateWorksheet = ExcelHandler.GetDuplicateWorksheet(yesterdayFile);
        var specimenIdCount = (ExcelHandler.GetDistinctSpecimenIdsFromWorksheet(duplicateWorksheet) ?? Array.Empty<string>()).Count();
        
        //Act
        var yesterdayDuplicateSpecimenList = ExcelHandler.GetYesterdayDuplicateSpecimenIds();
        
        //Assert
        yesterdayDuplicateSpecimenList.Should().HaveCount(specimenIdCount);
    }
    
    [Theory]
    [InlineData(new[]{"12345678","12345678"}, 1)]
    [InlineData(new[]{"1234567811","12345678554"}, 2)]
    [InlineData(new[]{"1234567811","12345678554","12345678554","1234567811"}, 2)]
    [InlineData(new[]{""}, 0)]
    [InlineData(new[]{"",""}, 0)]
    public void GetDistinctSpecimenIdsFromWorksheet_Should_Return_Correct_Counts(string[] specimenIds, int expectedCount)
    {
        //Arrange
        var currentRow = ExcelHandler.CreateWorksheet(ExcelHandler.DuplicatesWorksheetName, out _,
            out var worksheet);
        
        ExcelHandler.FormatWorksheet<Duplicates>(worksheet, currentRow);

        var ourDuplicatesList = specimenIds.Select(specimenId => new Duplicates { SpecId = specimenId }).ToList();
        
        foreach (var pendingSample in ourDuplicatesList)
        {
            currentRow++;
            ExcelHandler.AddWorksheetValues(worksheet, currentRow, pendingSample);
        }

        //Act
        var duplicateWorksheet = ExcelHandler.GetDistinctSpecimenIdsFromWorksheet(worksheet);

        //Assert
        duplicateWorksheet.Should().HaveCount(expectedCount);
        //here just asserting there is a header in the file
        //duplicateWorksheet.CellsUsed().Count().Should().BeGreaterThan(0);
    }
    
    [Fact]
    public void GetDuplicateWorksheet_Should_Return_A_IXLWorksheet()
    {
        //Arrange
        var yesterdayFile = ExcelHandler.GetLastDaysDuplicateFileInfo();

        //Act
        var duplicateWorksheet = ExcelHandler.GetDuplicateWorksheet(yesterdayFile);
        
        //Assert
        duplicateWorksheet.Should().NotBeNull();
        //here just asserting there is a header in the file
        duplicateWorksheet?.CellsUsed().Count().Should().BeGreaterThan(0);
    }
    
    [Fact]
    public void ExportDuplicatesToExcel_Should_Export_Records_To_Directory()
    {
        //Arrange
        var possibleDuplicates = DatabaseHandler.GetPossibleDuplicates().ToList();
        var fileDate = ExcelHandler.GetDateToAppendToFileName();
        var fileName = $"{ExcelHandler.BaseDirectory}Daily_Duplicates_{fileDate}_TestingFile.xlsx";

        //Act
        ExcelHandler.ExportDuplicatesToExcel(possibleDuplicates, fileName);
        
        //Assert
        File.Exists(fileName).Should().BeTrue();
        
        //cleanup
        File.Delete(fileName);

    }
    
    [Fact]
    public void ExportPendingSamplesToExcel_Should_Export_Records_To_Directory()
    {
        //Arrange
        var pendingSamplesList = DatabaseHandler.GetPendingSamples().ToList();
        var fileDate = ExcelHandler.GetDateToAppendToFileName();
        var fileName = $"{ExcelHandler.BaseDirectory}clinmicro_pending_list_{fileDate}.xlsx";

        //Act
        ExcelHandler.ExportPendingSamplesToExcel(pendingSamplesList, fileName);
        
        //Assert
        File.Exists(fileName).Should().BeTrue();
        
        //cleanup
        File.Delete(fileName);

    }

    [Theory]
    [ClassData(typeof(DuplicateTestData))]
    public void AddWorksheetValues_Duplicates_Should_Add_Record_To_Worksheet(Duplicates duplicateRowToAdd)
    {
        //Arrange
        var currentRow = ExcelHandler.CreateWorksheet(ExcelHandler.DuplicatesWorksheetName, out _,
            out var worksheet);
        GetDuplicateClassPropertyCount(out var propertyCount);
        
        //Act
        ExcelHandler.AddWorksheetValues(worksheet, currentRow,duplicateRowToAdd);
        
        //Assert
        worksheet.Should().NotBeNull();
        worksheet?.CellsUsed().Should().NotBeNullOrEmpty();
        worksheet?.Cells().Should().HaveCount(propertyCount);
        worksheet?.Row(1).Cell(1).GetString().Should().Be(duplicateRowToAdd.SpecId);
        worksheet?.Row(1).Cell(2).GetString().Should().Be(duplicateRowToAdd.ProcessInstName);
        worksheet?.Row(1).Cell(3).GetString().Should().Be(duplicateRowToAdd.ProcessDefName);
        worksheet?.Row(1).Cell(4).GetString().Should().Be(duplicateRowToAdd.ResevSuppPosition);
        worksheet?.Row(1).Cell(5).GetString().Should().Be(duplicateRowToAdd.ProcessingStatus);
        worksheet?.Row(1).Cell(6).GetString().Should().Be(duplicateRowToAdd.TestCodes);
        worksheet?.Row(1).Cell(7).GetString().Should().Be(duplicateRowToAdd.PendingRxns);
        worksheet?.Row(1).Cell(8).GetString().Should().Be(duplicateRowToAdd.RepeatReason);

    }
    
    public class DuplicateTestData:IEnumerable<object[]>
    {
        public IEnumerator<object[]> GetEnumerator()
        {
            yield return new object[] { new Duplicates
            {
                SpecId = "10293545310", 
                ProcessInstName = "230318_014139",
                ProcessDefName = "MagNa Pure 96 Extraction (ClinMicro)",
                ResevSuppPosition = "H2",
                ProcessingStatus = "MP96ExtrSuccess", 
                TestCodes = EmptyValue,
                PendingRxns = EmptyValue,
                RepeatReason = EmptyValue
            } };
            
            /*yield return new object[] { new Duplicates { Id = 2, FirstName = "Mary", LastName = null } };
            yield return new object[] { new Duplicates { Id = 3, FirstName = "Mary", LastName = null } };
            yield return new object[] { new Duplicates { Id = 4, FirstName = "", LastName = null } };
            yield return new object[] { new Duplicates { Id = 5, FirstName = "", LastName = "john" } };
            yield return new object[] { new Duplicates { Id = 6, FirstName = null, LastName = " " } };*/
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}