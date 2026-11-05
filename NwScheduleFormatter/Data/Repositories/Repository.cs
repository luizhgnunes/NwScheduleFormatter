using Microsoft.EntityFrameworkCore;

namespace NwScheduleFormatter.Data.Repositories;

/// <summary>
/// Generic implementation of the CRUD repository.
/// </summary>
/// <typeparam name="TEntity">Type of the entity.</typeparam>
public class Repository<TEntity> : IRepository<TEntity> where TEntity : class
{
    protected readonly NwScheduleFormatterDbContext _context;
    protected readonly DbSet<TEntity> _dbSet;

    public Repository(NwScheduleFormatterDbContext context)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _dbSet = context.Set<TEntity>();
    }

    public virtual async Task<IEnumerable<TEntity>> GetAllAsync()
    {
        return await _dbSet.ToListAsync();
    }

    public virtual async Task<TEntity?> GetByIdAsync(int id)
    {
        return await _dbSet.FindAsync(id);
    }

    public virtual async Task<TEntity> AddAsync(TEntity entity)
    {
        ArgumentNullException.ThrowIfNull(entity);
        var result = _dbSet.Add(entity);
        await SaveChangesAsync();
        return result.Entity;
    }

    public virtual async Task<TEntity> UpdateAsync(TEntity entity)
    {
        ArgumentNullException.ThrowIfNull(entity);
        var result = _dbSet.Update(entity);
        await SaveChangesAsync();
        return result.Entity;
    }

    public virtual async Task DeleteAsync(int id)
    {
        var entity = await GetByIdAsync(id);
        if (entity != null)
        {
            _dbSet.Remove(entity);
            await SaveChangesAsync();
        }
    }

    public virtual async Task<bool> ExistsAsync(int id)
    {
        return await _dbSet.FindAsync(id) != null;
    }

    public async Task SaveChangesAsync()
    {
        await _context.SaveChangesAsync();
    }
}
