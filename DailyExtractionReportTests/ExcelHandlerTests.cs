using DailyExtractionReports;
using DailyExtractionReports.Models;
using FluentAssertions;

namespace DailyExtractionReportTests;

public class ExcelHandlerTests
{
    
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
        var duplicatesType = typeof(Duplicates);
        var propertyCount = duplicatesType.Properties().Count();
        ExcelHandler.CreateWorksheet(worksheetName, out _, out var worksheet);
        
        //Act
        ExcelHandler.FormatWorksheet<Duplicates>(worksheet, currentRow);
        
        //Assert
        worksheet?.CellsUsed().Should().NotBeNullOrEmpty();
        worksheet?.CellsUsed().Should().HaveCount(propertyCount);
        worksheet?.CellsUsed().Select(c => c.Value.ToString()).Distinct().ToList()
            .Should().BeEquivalentTo(duplicatesType.Properties().Select(x => x.Name).ToList());
        

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
    
    /*[Theory]
    [InlineData(new[]{"1234567","1234567", ""}, 2)]
    public void AddDuplicatesToWorksheet_Should_Not_Contain_False_Duplicates_OrControls(string[] possibleDuplicates, int expectedNumberRows)
    {
        //Arrange
        const string worksheetName = "Test Worksheet";
        var currentRow = ExcelHandler.CreateWorksheet(worksheetName, out _, out var worksheet);

        //Act
        ExcelHandler.AddDuplicatesToWorksheet(new List<Duplicates>(), currentRow, worksheet);

        //Assert
        //duplicateWorksheet.Should().NotBeNull();
        //here just asserting there is a header in the file
        //duplicateWorksheet.CellsUsed().Count().Should().BeGreaterThan(0);
    }*/
}