namespace NwScheduleFormatter.Data.Repositories.Interfaces;

/// <summary>
/// Generic interface for CRUD repository operations.
/// </summary>
/// <typeparam name="TEntity">Type of the entity.</typeparam>
public interface IRepository<TEntity> where TEntity : class
{
    /// <summary>
    /// Gets all entities.
    /// </summary>
    Task<IEnumerable<TEntity>> GetAllAsync();

    /// <summary>
    /// Gets an entity by ID.
    /// </summary>
    Task<TEntity?> GetByIdAsync(int id);

    /// <summary>
    /// Adds a new entity.
    /// </summary>
    Task<TEntity> AddAsync(TEntity entity);

    /// <summary>
    /// Updates an existing entity.
    /// </summary>
    Task<TEntity> UpdateAsync(TEntity entity);

    /// <summary>
    /// Removes an entity.
    /// </summary>
    Task DeleteAsync(int id);

    /// <summary>
    /// Checks if an entity exists.
    /// </summary>
    Task<bool> ExistsAsync(int id);

    /// <summary>
    /// Saves changes to the database.
    /// </summary>
    Task SaveChangesAsync();
}