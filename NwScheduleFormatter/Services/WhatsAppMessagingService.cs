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
                var wamid = await SendMessageToContact(contact, meeting.TreasuresFromGodsWord.BibleReading, "N/A", meeting.Date, "3", "Salão Principal");
                await SaveSendingOfDesignation(meeting.Date, "3", meeting.TreasuresFromGodsWord.BibleReading, wamid);
            }
            if (!string.IsNullOrWhiteSpace(meeting.Apply1.Speaker))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.Apply1.Speaker, StringComparison.OrdinalIgnoreCase));
                var wamid = await SendMessageToContact(contact, meeting.Apply1.Speaker, meeting.Apply1.Assistant, meeting.Date, "4", "Salão Principal");
                await SaveSendingOfDesignation(meeting.Date, "4", meeting.Apply1.Speaker, wamid);
            }
            if (!string.IsNullOrWhiteSpace(meeting.Apply2.Speaker))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.Apply2.Speaker, StringComparison.OrdinalIgnoreCase));
                var wamid = await SendMessageToContact(contact, meeting.Apply2.Speaker, meeting.Apply2.Assistant, meeting.Date, "5", "Salão Principal");
                await SaveSendingOfDesignation(meeting.Date, "5", meeting.Apply2.Speaker, wamid);
            }
            if (!string.IsNullOrWhiteSpace(meeting.Apply3.Speaker))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.Apply3.Speaker, StringComparison.OrdinalIgnoreCase));
                var wamid = await SendMessageToContact(contact, meeting.Apply3.Speaker, meeting.Apply3.Assistant, meeting.Date, "6", "Salão Principal");
                await SaveSendingOfDesignation(meeting.Date, "6", meeting.Apply3.Speaker, wamid);
            }
            if (!string.IsNullOrWhiteSpace(meeting.Apply4.Speaker))
            {
                var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(meeting.Apply4.Speaker, StringComparison.OrdinalIgnoreCase));
                var wamid = await SendMessageToContact(contact, meeting.Apply4.Speaker, meeting.Apply4.Assistant, meeting.Date, "7", "Salão Principal");
                await SaveSendingOfDesignation(meeting.Date, "7", meeting.Apply4.Speaker, wamid);
            }
        }
    }

    private async Task<string?> SendMessageToContact(DesignationContact contact, string speakerName, string assistantName, DateOnly meetingDate, string partNumber, string location = "Salão Principal")
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

    private async Task SaveSendingOfDesignation(DateOnly meetingDate, string partNumber, string speakerName, string wamId)
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