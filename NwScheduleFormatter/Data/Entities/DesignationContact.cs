namespace NwScheduleFormatter.Data.Entities;

/// <summary>
/// Represents a designation contact.
/// Mapped from the tb_designation_contact table.
/// </summary>
public class DesignationContact
{
    /// <summary>
    /// Unique identifier.
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// Contact name.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// WhatsApp number for the contact.
    /// </summary>
    public string? WhatsAppNumber { get; set; }

    /// <summary>
    /// Indicates whether to skip sending messages to this contact.
    /// </summary>
    public bool SkipSendingMessage { get; set; }
}
