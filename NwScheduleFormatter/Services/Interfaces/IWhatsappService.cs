namespace NwScheduleFormatter.Services.Interfaces;

public interface IWhatsappService
{
    Task<string?> SendTextMessageAsync(string recipientPhoneNumber, string messageText);
    Task<string?> SendTemplateMessageAsync(string recipientPhoneNumber, string templateName, string languageCode = "pt_BR");
    Task<string?> SendTemplateMessageWithParametersAsync(string recipientPhoneNumber, string templateName, string[] parameters, string languageCode = "pt_BR");
    Task<string?> SendImageMessageAsync(string recipientPhoneNumber, string imageUrl, string caption = null);
    Task<string?> SendDocumentMessageAsync(string recipientPhoneNumber, string documentUrl, string filename = null);
    Task<bool> MarkMessageAsReadAsync(string messageId);
    Task<string> GetMessageTemplatesAsync();
    Task<bool> CreateMessageTemplateAsync(string templateName, string templateText, string category = "TRANSACTIONAL", string language = "pt_BR");
    Task<string> GetPhoneNumberDetailsAsync();
}
