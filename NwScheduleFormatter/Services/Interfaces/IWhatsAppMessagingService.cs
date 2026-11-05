using NwScheduleFormatter.Models;

namespace NwScheduleFormatter.Services.Interfaces;

public interface IWhatsAppMessagingService
{
    /// <summary>
    /// Send confirmation requests for the provided meetings via WhatsApp.
    /// </summary>
    /// <param name="meetings"><see cref="List{Meeting}"/></param>
    /// <returns></returns>
    Task SendConfirmationRequestsAsync(List<Meeting> meetings);
}
