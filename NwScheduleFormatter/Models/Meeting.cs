namespace NwScheduleFormatter.Models;

internal class Meeting
{
    public Meeting() { }

    public Meeting(DateOnly date)
    {
        Date = date;
    }

    public DateOnly Date { get; set; }

    public string OpeningPrayer { get; set; }
    public string Chairman { get; set; } // President
    public Song InitialSong { get; set; } = new();

    public TreasuresFromGodsWord TreasuresFromGodsWord { get; set; } = new();

    public Presentation Apply1 { get; set; } = new();
    public Presentation Apply2 { get; set; } = new();
    public Presentation Apply3 { get; set; } = new();
    public Presentation Apply4 { get; set; } = new();

    public Song MiddleSong { get; set; } = new();

    public Speech LivingAsChristhians1 { get; set; } = new();
    public Speech LivingAsChristhians2 { get; set; } = new();
    public CongregationBibleStudy CongregationBibleStudy { get; set; } = new();

    public Song FinalSong { get; set; } = new();
    public string ClosingPrayer { get; set; }
}

