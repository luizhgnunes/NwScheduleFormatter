namespace NwScheduleFormatter.Data.Entities;

/// <summary>
/// Represents the WhatsApp configuration settings.
/// Mapped from the tb_whatsapp_config table.
/// </summary>
public class WhatsAppConfig
{
    /// <summary>
    /// Unique identifier.
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// WhatsApp API access token.
    /// </summary>
    public string? AccessToken { get; set; }

    /// <summary>
    /// WhatsApp phone number ID.
    /// </summary>
    public string? PhoneNumberId { get; set; }

    /// <summary>
    /// WhatsApp webhook URL.
    /// </summary>
    public string? WebhookUrl { get; set; }
}
