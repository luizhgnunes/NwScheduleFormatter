using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using NwScheduleFormatter.Models;

namespace NwScheduleFormatter.Data;

public class SuperintendentVisitSchedule
{
    private static Dictionary<DateOnly, SuperintendentVisitSpeech> SuperintendentVisit { get; set; } = [];

    public void LoadFromJson(string jsonFilePath)
    {
        if (!File.Exists(jsonFilePath))
            return;

        try
        {
            var json = File.ReadAllText(jsonFilePath);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                Converters = { new JsonStringEnumConverter() }
            };

            var visits = JsonSerializer.Deserialize<List<SuperintendentVisitData>>(json, options);

            if (visits != null)
            {
                foreach (var visit in visits)
                {
                    if (DateOnly.TryParse(visit.Date, out var date))
                    {
                        SuperintendentVisit.Add(
                            date, 
                            new SuperintendentVisitSpeech { Speaker = visit.Speaker, Theme = visit.Theme, Song = visit.Song }
                        );
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error while loading superintendent visits JSON file: {ex.Message}");
        }
    }

    public bool IsSuperintendentVisit(DateOnly date)
    {
        return SuperintendentVisit.Any(s => date >= s.Key && date < s.Key.AddDays(7));
    }

    public SuperintendentVisitSpeech GetSuperintendentVisitSpeech(DateOnly mondayDate)
    {
        return SuperintendentVisit[mondayDate];
    }

    private class SuperintendentVisitData
    {
        [JsonPropertyName("date")]
        public string Date { get; set; }

        [JsonPropertyName("speaker")]
        public string Speaker { get; set; }

        [JsonPropertyName("theme")]
        public string Theme { get; set; }

        [JsonPropertyName("song")]
        public Song Song { get; set; }
    }
}
