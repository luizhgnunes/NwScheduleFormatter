namespace NwScheduleFormatter.Data.Entities;

using NwScheduleFormatter.Data.Enums;

/// <summary>
/// Represents a designation sending via WhatsApp.
/// Mapped from the tb_sending_of_designation table.
/// </summary>
public class SendingOfDesignation
{
    /// <summary>
    /// Unique identifier.
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// Unique reference for the sending (10 characters).
    /// </summary>
    public string? Reference { get; set; }

    /// <summary>
    /// Name of the designation.
    /// </summary>
    public string? DesignatedName { get; set; }

    /// <summary>
    /// Date of the designation.
    /// </summary>
    public DateTime DesignationDate { get; set; }

    /// <summary>
    /// Record creation date.
    /// </summary>
    public DateTime CreationDate { get; set; }

    /// <summary>
    /// Indicates whether the sending was canceled.
    /// </summary>
    public bool IsCanceled { get; set; }

    /// <summary>
    /// WhatsApp message ID (unique identifier of the message).
    /// </summary>
    public string? WamId { get; set; }

    /// <summary>
    /// Date when the message was sent.
    /// </summary>
    public DateTime? SendDate { get; set; }

    /// <summary>
    /// Indicates whether the sending was confirmed.
    /// </summary>
    public bool IsConfirmed { get; set; }

    /// <summary>
    /// Confirmation date.
    /// </summary>
    public DateTime? ConfirmationDate { get; set; }

    /// <summary>
    /// Source of the confirmation (NotConfirmed, WhatsApp, or Wpf).
    /// </summary>
    public ConfirmationSource ConfirmationSource { get; set; }

    /// <summary>
    /// Error message if sending failed.
    /// </summary>
    public string? SendingError { get; set; }
}
