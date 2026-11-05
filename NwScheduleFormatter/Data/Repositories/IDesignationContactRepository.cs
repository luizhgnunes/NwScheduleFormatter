using NwScheduleFormatter.Data.Entities;

namespace NwScheduleFormatter.Data.Repositories;

/// <summary>
/// Specific interface for operations with DesignationContact.
/// </summary>
public interface IDesignationContactRepository : IRepository<DesignationContact>
{
    /// <summary>
    /// Gets all contacts that should receive messages.
    /// </summary>
    Task<IEnumerable<DesignationContact>> GetActiveContactsAsync();
}
