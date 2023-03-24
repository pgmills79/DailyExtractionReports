using DailyExtractionReports;
using DailyExtractionReports.Models;
using FluentAssertions;

namespace DailyExtractionReportTests;

public class DatabaseHandlerTests
{
    
    [Fact]
    public void GetPossibleDuplicates_Should_Return_Records()
    {
        //Arrange

        //Act
        var result = DatabaseHandler.GetPossibleDuplicates().ToList();
        
        //Assert
        result.Should().NotBeNullOrEmpty();
        result.Should().HaveCountGreaterThan(0);

    }
    
    [Fact]
    public void GetPendingSamples_Should_Return_Records()
    {
        //Arrange

        //Act
        var result = DatabaseHandler.GetPendingSamples().ToList();
        
        //Assert
        result.Should().NotBeNullOrEmpty();
        result.Should().HaveCountGreaterThan(0);

    }
    
    [Fact]
    public void GetConnectionString_Should_Return_ConnectionString()
    {
        //Arrange

        //Act
        var result = DatabaseHandler.GetConnectionString();
        
        //Assert
        result.Should().NotBeNullOrEmpty();
        result.Should().Contain("Data Source=");
        result.Should().Contain("Initial Catalog=");
        result.Should().Contain("Integrated Security=SSPI");
        //result.Should().HaveCountGreaterThan(0);

    }
    
}