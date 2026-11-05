using NwScheduleFormatter.Models;

namespace NwScheduleFormatter.Data.Interfaces;

public interface IJwMeetingReader
{
    Task<int> CountMonthMeetings(int year, int month);
    Task<JwWebsiteMeeting> ReadFromJwWebsiteAsync(DateOnly mondayDate);
}
