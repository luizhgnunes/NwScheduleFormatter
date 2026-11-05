using NwScheduleFormatter.Data.Entities;

namespace NwScheduleFormatter.Data.Repositories;

/// <summary>
/// Specific interface for operations with WhatsAppConfig.
/// </summary>
public interface IWhatsAppConfigRepository : IRepository<WhatsAppConfig>
{
    /// <summary>
    /// Gets the first (and typically only) WhatsApp configuration.
    /// </summary>
    Task<WhatsAppConfig?> GetConfigAsync();
}
