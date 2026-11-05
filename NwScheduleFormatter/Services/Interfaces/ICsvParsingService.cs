using NwScheduleFormatter.Models;

namespace NwScheduleFormatter.Services.Interfaces;

/// <summary>
/// Interface for CSV parsing services.
/// </summary>
public interface ICsvParsingService
{
    /// <summary>
    /// Reads and parses meeting schedule from a CSV file.
    /// </summary>
    Task<List<Meeting>> GetMeetingScheduleFromCsvAsync(string csvFilePath, int year, int month, IProgress<int> progress);
}
