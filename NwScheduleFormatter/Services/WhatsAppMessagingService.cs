using NwScheduleFormatter.Data.Entities;
using NwScheduleFormatter.Data.Repositories.Interfaces;
using NwScheduleFormatter.Models;
using NwScheduleFormatter.Services.Interfaces;

namespace NwScheduleFormatter.Services;

public class WhatsAppMessagingService : IWhatsAppMessagingService
{
    private const string NEW_DESIGNATION_TEMPLATE_NAME = "nova_designacao";

    private readonly IWhatsappService _whatsappService;
    private readonly IDesignationContactRepository _designationContactRepository;
    private readonly ISendingOfDesignationRepository _sendingOfDesignationRepository;

    public WhatsAppMessagingService(IWhatsappService whatsappService, IDesignationContactRepository designationContactRepository, ISendingOfDesignationRepository sendingOfDesignationRepository)
    {
        _whatsappService = whatsappService;
        _designationContactRepository = designationContactRepository;
        _sendingOfDesignationRepository = sendingOfDesignationRepository;
    }

    public async Task SendConfirmationRequestsAsync(List<Meeting> meetings)
    {
        var activeContacts = await _designationContactRepository.GetActiveContactsAsync();
        foreach (var meeting in meetings)
        {
            if (!string.IsNullOrWhiteSpace(meeting.TreasuresFromGodsWord.BibleReading))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.TreasuresFromGodsWord.BibleReading, StringComparison.OrdinalIgnoreCase));
                var wamid = await SendMessageToContactAsync(contact, meeting.TreasuresFromGodsWord.BibleReading, "N/A", meeting.Date, "3", "Salão Principal");
                await SaveSendingOfDesignationAsync(meeting.Date, "3", meeting.TreasuresFromGodsWord.BibleReading, wamid);
            }
            if (!string.IsNullOrWhiteSpace(meeting.Apply1.Speaker))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.Apply1.Speaker, StringComparison.OrdinalIgnoreCase));
                var assistant = string.IsNullOrWhiteSpace(meeting.Apply1.Assistant) ? "N/A" : meeting.Apply1.Assistant;
                var wamid = await SendMessageToContactAsync(contact, meeting.Apply1.Speaker, assistant, meeting.Date, "4", "Salão Principal");
                await SaveSendingOfDesignationAsync(meeting.Date, "4", meeting.Apply1.Speaker, wamid);
            }
            if (!string.IsNullOrWhiteSpace(meeting.Apply2.Speaker))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.Apply2.Speaker, StringComparison.OrdinalIgnoreCase));
                var assistant = string.IsNullOrWhiteSpace(meeting.Apply2.Assistant) ? "N/A" : meeting.Apply2.Assistant;
                var wamid = await SendMessageToContactAsync(contact, meeting.Apply2.Speaker, assistant, meeting.Date, "5", "Salão Principal");
                await SaveSendingOfDesignationAsync(meeting.Date, "5", meeting.Apply2.Speaker, wamid);
            }
            if (!string.IsNullOrWhiteSpace(meeting.Apply3.Speaker))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.Apply3.Speaker, StringComparison.OrdinalIgnoreCase));
                var assistant = string.IsNullOrWhiteSpace(meeting.Apply3.Assistant) ? "N/A" : meeting.Apply3.Assistant;
                var wamid = await SendMessageToContactAsync(contact, meeting.Apply3.Speaker, assistant, meeting.Date, "6", "Salão Principal");
                await SaveSendingOfDesignationAsync(meeting.Date, "6", meeting.Apply3.Speaker, wamid);
            }
            if (!string.IsNullOrWhiteSpace(meeting.Apply4.Speaker))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.Apply4.Speaker, StringComparison.OrdinalIgnoreCase));
                var assistant = string.IsNullOrWhiteSpace(meeting.Apply4.Assistant) ? "N/A" : meeting.Apply4.Assistant;
                var wamid = await SendMessageToContactAsync(contact, meeting.Apply4.Speaker, assistant, meeting.Date, "7", "Salão Principal");
                await SaveSendingOfDesignationAsync(meeting.Date, "7", meeting.Apply4.Speaker, wamid);
            }
        }
    }

    private async Task<string?> SendMessageToContactAsync(DesignationContact contact, string speakerName, string assistantName, DateOnly meetingDate, string partNumber, string location = "Salão Principal")
    {
        if (contact != null && !string.IsNullOrWhiteSpace(contact.WhatsAppNumber))
        {
            return await _whatsappService.SendTemplateMessageWithParametersAsync(
                contact.WhatsAppNumber,
                NEW_DESIGNATION_TEMPLATE_NAME,
                [speakerName, assistantName, meetingDate.ToString("dd/MM/yyyy"), partNumber, location]);
        }
        return null;
    }

    private async Task SaveSendingOfDesignationAsync(DateOnly meetingDate, string partNumber, string speakerName, string wamId)
    {
        var sending = new SendingOfDesignation
        {
            Reference = $"{meetingDate.ToString("yyyyMMdd")}-{partNumber}",
            DesignatedName = speakerName,
            DesignationDate = meetingDate.ToDateTime(TimeOnly.MinValue),
            CreationDate = DateTime.Now,
            IsCanceled = false,
            WamId = wamId,
            SendDate = DateTime.Now,
            IsConfirmed = false,
            ConfirmationDate = null
        };

        await _sendingOfDesignationRepository.AddAsync(sending);
    }
}