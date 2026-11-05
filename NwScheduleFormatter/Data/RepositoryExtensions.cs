using Microsoft.Extensions.DependencyInjection;
using NwScheduleFormatter.Data.Repositories;

namespace NwScheduleFormatter.Data;

/// <summary>
/// Extensions to facilitate access to application repositories.
/// </summary>
public static class RepositoryExtensions
{
    /// <summary>
    /// Gets the SendingOfDesignation repository from the ServiceProvider.
    /// </summary>
    public static ISendingOfDesignationRepository GetSendingOfDesignationRepository()
    {
        var serviceProvider = (App.ServiceProvider as ServiceProvider);
        if (serviceProvider == null)
            throw new InvalidOperationException("ServiceProvider was not initialized.");

        return serviceProvider.GetRequiredService<ISendingOfDesignationRepository>();
    }
}
