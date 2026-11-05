using Microsoft.EntityFrameworkCore;
using NwScheduleFormatter.Data.Entities;

namespace NwScheduleFormatter.Data.Repositories;

/// <summary>
/// Specific repository implementation for WhatsAppConfig.
/// </summary>
public class WhatsAppConfigRepository : Repository<WhatsAppConfig>, IWhatsAppConfigRepository
{
    public WhatsAppConfigRepository(NwScheduleFormatterDbContext context)
        : base(context)
    {
    }

    public async Task<WhatsAppConfig?> GetConfigAsync()
    {
        return await _dbSet.FirstOrDefaultAsync();
    }
}
