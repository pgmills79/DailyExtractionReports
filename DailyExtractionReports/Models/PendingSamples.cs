namespace DailyExtractionReports.Models;

public class PendingSamples
{
   
    public string SpecId { get; set; } = null!;
    public string Title { get; set; } = null!;
    public string Batch { get; set; } = null!;
    public string TestsOrdered { get; set; } = null!;
    public string JanusA { get; set; } = null!;
    public string JanusB { get; set; } = null!;
    public string JanusC { get; set; } = null!;
    public string Analysis { get; set; } = null!;
    public string Verifying { get; set; } = null!;
    public string TestsPending { get; set; } = null!;
    public string AllCompletePendingBatches { get; set; } = null!;
}