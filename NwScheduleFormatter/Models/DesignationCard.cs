namespace NwScheduleFormatter.Models;

public class DesignationCard
{
    public DateOnly Date { get; set; }
    public Presentation Presentation { get; set; }
    public short PartNumber { get; set; }

    public DesignationCard() { }
    public DesignationCard(DateOnly date, Presentation presentation, short partNumber)
    {
        Date = date;
        Presentation = presentation;
        PartNumber = partNumber;
    }
}
