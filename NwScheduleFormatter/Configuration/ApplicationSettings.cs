using NwScheduleFormatter.Enumerations;

namespace NwScheduleFormatter.Configuration;

public class ApplicationSettings
{
    public const string CONGREGATION_NAME = "EURICO CHAVES";
    public const string START_TIME = "19:30";
    public const string DEFAULT_OUTPUT_DIRECTORY = @"C:\NwScheduleFormatter\Files";
    public static MeetingDay MeetingDay => MeetingDay.Wednesday;
}