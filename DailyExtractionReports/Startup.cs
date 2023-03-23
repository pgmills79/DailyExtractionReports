using Microsoft.Extensions.Configuration;

namespace DailyExtractionReports;

public class Startup
{
    private IConfigurationBuilder? Builder { get; }
    public IConfiguration? Config { get; }
    public Startup()
    {
        Builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false);
        
        Config = Builder?.Build();
    }
}