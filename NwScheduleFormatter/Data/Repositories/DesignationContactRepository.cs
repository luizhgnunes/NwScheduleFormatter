using Microsoft.EntityFrameworkCore;
using NwScheduleFormatter.Data.Entities;
using NwScheduleFormatter.Data.Repositories.Interfaces;

namespace NwScheduleFormatter.Data.Repositories;

/// <summary>
/// Specific repository implementation for DesignationContact.
/// </summary>
public class DesignationContactRepository : Repository<DesignationContact>, IDesignationContactRepository
{
    public DesignationContactRepository(NwScheduleFormatterDbContext context)
        : base(context)
    {
    }

    public async Task<IEnumerable<DesignationContact>> GetActiveContactsAsync()
    {
        return await _dbSet.Where(x => !x.SkipSendingMessage).ToListAsync();
    }
}