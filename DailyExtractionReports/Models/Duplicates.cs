namespace DailyExtractionReports.Models;

public class Duplicates
{
   
    public string SpecId { get; set; } = null!;
    public string ProcessInstName { get; set; } = null!;
    public string ProcessDefName { get; set; } = null!;
    public string ResevSuppPosition { get; set; } = null!;
    public string ProcessingStatus { get; set; } = null!;
    public string TestCodes { get; set; } = null!;
    public string PendingRxns { get; set; } = null!;
    public string RepeatReason { get; set; } = null!;
}