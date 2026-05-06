using NwScheduleFormatter.Models;

namespace NwScheduleFormatter.Data;

public class SuperintendentVisitSchedule
{
    private Dictionary<DateOnly, SuperintendentVisitSpeech> SuperintendentVisit { get; set; } = [];

    public SuperintendentVisitSchedule()
    {
        SuperintendentVisit.Add(
            new DateOnly(2026, 6, 1), new SuperintendentVisitSpeech
            {
                Speaker = "Nome do Superintendente",
                Theme = "Discurso do Superintendente de Circuito"
            });
    }

    public bool IsSuperintendentVisit(DateOnly date)
    {
        return SuperintendentVisit.Any(s => date >= s.Key && date < s.Key.AddDays(7));
    }

    public SuperintendentVisitSpeech GetSuperintendentVisitSpeech(DateOnly mondayDate)
    {
        return SuperintendentVisit[mondayDate];
    }
}
