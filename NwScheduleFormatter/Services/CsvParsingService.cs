using NwScheduleFormatter.Data;
using NwScheduleFormatter.Data.Interfaces;
using NwScheduleFormatter.Enumerations;
using NwScheduleFormatter.Models;
using NwScheduleFormatter.Services.Interfaces;

namespace NwScheduleFormatter.Services;

/// <summary>
/// Service for parsing meeting schedules from CSV files.
/// </summary>
public class CsvParsingService : ICsvParsingService
{
    private readonly IJwMeetingReader _jwMeetingReader;

    public CsvParsingService(IJwMeetingReader jwMeetingReader)
    {
        _jwMeetingReader = jwMeetingReader;
    }

    public async Task<List<Meeting>> GetMeetingScheduleFromCsvAsync(string csvFilePath, int year, int month, IProgress<int> progress)
    {
        progress?.Report(0);
        var nwCsvRecords = NwCsvReader.ReadFromCsv(csvFilePath, year, month);

        var meetings = new List<Meeting>();
        var totalMeetings = await _jwMeetingReader.CountMonthMeetings(year, month);
        
        foreach (var record in nwCsvRecords)
        {
            var meeting = meetings.Find(x => x.Date == record.Date);
            if (meeting == null)
            {
                meeting = new Meeting(record.Date);

                var jwWebsiteMeeting = await _jwMeetingReader.ReadFromJwWebsiteAsync(meeting.Date);
                meeting.WeeklyBibleReading = jwWebsiteMeeting.WeeklyBibleReading;
                meeting.InitialSong = jwWebsiteMeeting.InitialSong;
                meeting.MiddleSong = jwWebsiteMeeting.MiddleSong;
                meeting.FinalSong = jwWebsiteMeeting.FinalSong;
                meeting.Apply1.DurationMinutes = jwWebsiteMeeting.Apply1DurationMinutes;
                meeting.Apply2.DurationMinutes = jwWebsiteMeeting.Apply2DurationMinutes;
                meeting.Apply3.DurationMinutes = jwWebsiteMeeting.Apply3DurationMinutes;
                meeting.Apply4.DurationMinutes = jwWebsiteMeeting.Apply4DurationMinutes;
                meeting.LivingAsChristhians1.DurationMinutes = jwWebsiteMeeting.LivingAsChristhians1DurationMinutes;
                meeting.LivingAsChristhians2.DurationMinutes = jwWebsiteMeeting.LivingAsChristhians2DurationMinutes;

                meetings.Add(meeting);
            }

            var partType = Enum.Parse<PartType>(record.PartType);
            switch (partType)
            {
                case PartType.Chairman:
                    meeting.Chairman = record.Person;
                    break;
                case PartType.OpeningPrayer:
                    meeting.OpeningPrayer = record.Person;
                    break;
                case PartType.ClosingPrayer:
                    meeting.ClosingPrayer = record.Person;
                    break;

                case PartType.TreasuresTalk:
                    meeting.TreasuresFromGodsWord.Treasures
                            = new Speech
                            {
                                Theme = record.Assignment,
                                Speaker = record.Person,
                                DurationMinutes = 10
                            };
                    break;
                case PartType.SpiritualGems:
                    meeting.TreasuresFromGodsWord.SpiritualGems = record.Person;
                    break;
                case PartType.BibleReading:
                    meeting.TreasuresFromGodsWord.BibleReading = record.Person;
                    break;

                case PartType.Apply1:
                    meeting.Apply1.Theme = record.Assignment;
                    meeting.Apply1.Speaker = record.Person;
                    break;
                case PartType.Apply2:
                    meeting.Apply2.Theme = record.Assignment;
                    meeting.Apply2.Speaker = record.Person;
                    break;
                case PartType.Apply3:
                    meeting.Apply3.Theme = record.Assignment;
                    meeting.Apply3.Speaker = record.Person;
                    break;
                case PartType.Apply4:
                    meeting.Apply4.Theme = record.Assignment;
                    meeting.Apply4.Speaker = record.Person;
                    break;
                case PartType.Apply1Assistant:
                    meeting.Apply1.Assistant = record.Person;
                    break;
                case PartType.Apply2Assistant:
                    meeting.Apply2.Assistant = record.Person;
                    break;
                case PartType.Apply3Assistant:
                    meeting.Apply3.Assistant = record.Person;
                    break;
                case PartType.Apply4Assistant:
                    meeting.Apply4.Assistant = record.Person;
                    break;

                case PartType.Living1:
                    meeting.LivingAsChristhians1.Theme = record.Assignment;
                    meeting.LivingAsChristhians1.Speaker = record.Person;
                    break;
                case PartType.Living2:
                    meeting.LivingAsChristhians2.Theme = record.Assignment;
                    meeting.LivingAsChristhians2.Speaker = record.Person;
                    break;
                case PartType.CBS:
                    meeting.CongregationBibleStudy.Speaker = record.Person;
                    break;
                case PartType.CBSReader:
                    meeting.CongregationBibleStudy.Reader = record.Person;
                    break;
            }
            var progressValue = (int)((meetings.Count / (double)totalMeetings) * 80);
            progress?.Report(progressValue);
        }

        progress?.Report(80);

        return meetings;
    }
}
