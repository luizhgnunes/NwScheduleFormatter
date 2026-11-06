using Microsoft.EntityFrameworkCore;
using NwScheduleFormatter.Data.Entities;
using NwScheduleFormatter.Data.Repositories.Interfaces;

namespace NwScheduleFormatter.Data.Repositories;

/// <summary>
/// Specific repository implementation for SendingOfDesignation.
/// </summary>
public class SendingOfDesignationRepository : Repository<SendingOfDesignation>, ISendingOfDesignationRepository
{
    public SendingOfDesignationRepository(NwScheduleFormatterDbContext context)
        : base(context)
    {
    }

    public async Task<SendingOfDesignation?> GetSentDesignationAsync(string reference, string name)
    {
        ArgumentNullException.ThrowIfNullOrEmpty(reference);
        return await _dbSet.FirstOrDefaultAsync(x => x.Reference == reference && x.DesignatedName == name);
    }

    public async Task<IEnumerable<SendingOfDesignation>> GetActiveSendingsAsync()
    {
        return await _dbSet.Where(x => !x.IsCanceled).ToListAsync();
    }

    public async Task<IEnumerable<SendingOfDesignation>> GetConfirmedSendingsAsync()
    {
        return await _dbSet.Where(x => x.IsConfirmed).ToListAsync();
    }

    public async Task<SendingOfDesignation?> GetByWamIdAsync(string wamId)
    {
        ArgumentNullException.ThrowIfNullOrEmpty(wamId);
        return await _dbSet.FirstOrDefaultAsync(x => x.WamId == wamId);
    }
}