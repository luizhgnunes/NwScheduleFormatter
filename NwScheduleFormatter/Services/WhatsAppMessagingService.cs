using NwScheduleFormatter.Configuration;
using NwScheduleFormatter.Data.Entities;
using NwScheduleFormatter.Data.Repositories.Interfaces;
using NwScheduleFormatter.Enumerations;
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
                await SendMessageAndSaveAsync(activeContacts, "3", meeting.Date, meeting.TreasuresFromGodsWord.BibleReading);
            
            if (!string.IsNullOrWhiteSpace(meeting.Apply1.Speaker))
                await SendMessageAndSaveAsync(activeContacts, "4", meeting.Date, meeting.Apply1.Speaker, meeting.Apply1.Assistant);
            
            if (!string.IsNullOrWhiteSpace(meeting.Apply2.Speaker))
                await SendMessageAndSaveAsync(activeContacts, "5", meeting.Date, meeting.Apply2.Speaker, meeting.Apply2.Assistant);

            if (!string.IsNullOrWhiteSpace(meeting.Apply3.Speaker))
                await SendMessageAndSaveAsync(activeContacts, "6", meeting.Date, meeting.Apply3.Speaker, meeting.Apply3.Assistant);

            if (!string.IsNullOrWhiteSpace(meeting.Apply4.Speaker))
                await SendMessageAndSaveAsync(activeContacts, "7", meeting.Date, meeting.Apply4.Speaker, meeting.Apply4.Assistant);
        }
    }

    private async Task SendMessageAndSaveAsync(IEnumerable<DesignationContact> activeContacts, string partNumber, DateOnly meetingDate, string speakerName, string? assistant = null)
    {
        var reference = $"{meetingDate.ToString("yyyyMMdd")}-{partNumber}";
        var designationDate = meetingDate.AddDays((int)ApplicationSettings.MeetingDay);
        if (await CheckIfConfirmationRequestWasSentAsync(reference, speakerName))
            return;
        
        var contact = activeContacts.FirstOrDefault(x => x.Name.Equals(speakerName, StringComparison.OrdinalIgnoreCase));
        assistant = string.IsNullOrWhiteSpace(assistant) ? "N/A" : assistant;
        var wamid = await SendMessageToContactAsync(contact, speakerName, assistant, designationDate, partNumber, "Salão Principal");
        await SaveSendingOfDesignationAsync(reference, speakerName, designationDate, wamid);
    }

    private async Task<bool> CheckIfConfirmationRequestWasSentAsync(string reference, string name)
    {
        var designation = await _sendingOfDesignationRepository.GetSentDesignationAsync(reference, name);
        return designation != null;
    }

    private async Task<string?> SendMessageToContactAsync(DesignationContact contact, string speakerName, string assistantName, DateOnly designationDate, string partNumber, string location = "Salão Principal")
    {
        if (contact != null && !string.IsNullOrWhiteSpace(contact.WhatsAppNumber))
        {
            return await _whatsappService.SendTemplateMessageWithParametersAsync(
                contact.WhatsAppNumber,
                NEW_DESIGNATION_TEMPLATE_NAME,
                [speakerName, assistantName, designationDate.ToString("dd/MM/yyyy"), partNumber, location]);
        }
        return null;
    }

    private async Task SaveSendingOfDesignationAsync(string reference, string speakerName, DateOnly designationDate, string wamId)
    {
        var sending = new SendingOfDesignation
        {
            Reference = reference,
            DesignatedName = speakerName,
            DesignationDate = designationDate.ToDateTime(TimeOnly.MinValue),
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