namespace NwScheduleFormatter.Data.Enums;

/// <summary>
/// Represents the source of a designation sending confirmation.
/// </summary>
public enum ConfirmationSource : byte
{
    /// <summary>
    /// Not yet confirmed.
    /// </summary>
    NotConfirmed = 0,

    /// <summary>
    /// Confirmed via WhatsApp.
    /// </summary>
    WhatsApp = 1,

    /// <summary>
    /// Confirmed via WPF application.
    /// </summary>
    Wpf = 2
}
