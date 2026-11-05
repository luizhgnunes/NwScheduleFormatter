using NwScheduleFormatter.Data.Entities;

namespace NwScheduleFormatter.Data.Repositories.Interfaces;

/// <summary>
/// Specific interface for operations with SendingOfDesignation.
/// </summary>
public interface ISendingOfDesignationRepository : IRepository<SendingOfDesignation>
{
    /// <summary>
    /// Gets a sending of designation by its unique reference.
    /// </summary>
    Task<SendingOfDesignation?> GetByReferenceAsync(string reference);

    /// <summary>
    /// Gets all non-canceled sendings.
    /// </summary>
    Task<IEnumerable<SendingOfDesignation>> GetActiveSendingsAsync();

    /// <summary>
    /// Gets all confirmed sendings.
    /// </summary>
    Task<IEnumerable<SendingOfDesignation>> GetConfirmedSendingsAsync();

    /// <summary>
    /// Gets a sending of designation by WhatsApp message ID.
    /// </summary>
    Task<SendingOfDesignation?> GetByWamIdAsync(string wamId);
}