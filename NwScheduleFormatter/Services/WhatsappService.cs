using NwScheduleFormatter.Services.Interfaces;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NwScheduleFormatter.Services;

public class WhatsappService : IWhatsappService
{
    private string _accessToken;
    private string _phoneNumberId;
    private readonly string _businessAccountId;
    private readonly HttpClient _httpClient;
    private const string BaseUrl = "https://graph.facebook.com/v20.0";

    public WhatsappService(string businessAccountId = "")
    {
        _businessAccountId = businessAccountId ?? throw new ArgumentNullException(nameof(businessAccountId));
        _httpClient = new HttpClient();
        LoadWhatsappCredentials();
    }

    private void LoadWhatsappCredentials(string jsonFilePath = "Configuration/whatsapp-config.json")
    {
        string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, jsonFilePath);

        if (!File.Exists(configPath))
            return;

        try
        {
            var json = File.ReadAllText(configPath);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };

            var config = JsonSerializer.Deserialize<WhatsappConfig>(json, options);

            if (config?.Whatsapp != null)
            {
                _accessToken = config.Whatsapp.AccessToken;
                _phoneNumberId = config.Whatsapp.PhoneNumberId;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error while loading WhatsApp configuration JSON file: {ex.Message}");
        }
    }

    /// <summary>
    /// Sends a text message via WhatsApp Cloud API
    /// </summary>
    public async Task<string?> SendTextMessageAsync(string recipientPhoneNumber, string messageText)
    {
        try
        {
            var requestBody = new
            {
                messaging_product = "whatsapp",
                to = recipientPhoneNumber,
                type = "text",
                text = new { body = messageText }
            };

            return await SendMessageAsync(requestBody);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending text message: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Sends a template message via WhatsApp Cloud API
    /// </summary>
    public async Task<string?> SendTemplateMessageAsync(string recipientPhoneNumber, string templateName, string languageCode = "pt_BR")
    {
        try
        {
            var requestBody = new
            {
                messaging_product = "whatsapp",
                to = recipientPhoneNumber,
                type = "template",
                template = new
                {
                    name = templateName,
                    language = new { code = languageCode }
                }
            };

            return await SendMessageAsync(requestBody);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending template message: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Sends a template message with parameters via WhatsApp Cloud API
    /// </summary>
    public async Task<string?> SendTemplateMessageWithParametersAsync(
        string recipientPhoneNumber,
        string templateName,
        string[] parameters,
        string languageCode = "pt_BR")
    {
        try
        {
            var requestBody = new
            {
                messaging_product = "whatsapp",
                to = recipientPhoneNumber,
                type = "template",
                template = new
                {
                    name = templateName,
                    language = new { code = languageCode },
                    components = new[]
                    {
                        new
                        {
                            type = "body",
                            parameters = ConvertParametersToList(parameters)
                        }
                    }
                }
            };

            return await SendMessageAsync(requestBody);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending template message with parameters: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Sends a message with an image
    /// </summary>
    public async Task<string?> SendImageMessageAsync(string recipientPhoneNumber, string imageUrl, string caption = null)
    {
        try
        {
            var requestBody = new
            {
                messaging_product = "whatsapp",
                to = recipientPhoneNumber,
                type = "image",
                image = new
                {
                    link = imageUrl,
                    caption = caption
                }
            };

            return await SendMessageAsync(requestBody);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending image message: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Sends a message with a document
    /// </summary>
    public async Task<string?> SendDocumentMessageAsync(string recipientPhoneNumber, string documentUrl, string filename = null)
    {
        try
        {
            var requestBody = new
            {
                messaging_product = "whatsapp",
                to = recipientPhoneNumber,
                type = "document",
                document = new
                {
                    link = documentUrl,
                    filename = filename
                }
            };

            return await SendMessageAsync(requestBody);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending document message: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Marks a message as read
    /// </summary>
    public async Task<bool> MarkMessageAsReadAsync(string messageId)
    {
        try
        {
            var requestBody = new
            {
                messaging_product = "whatsapp",
                status = "read",
                message_id = messageId
            };

            var url = $"{BaseUrl}/{_phoneNumberId}/messages?access_token={_accessToken}";
            var content = new StringContent(
                JsonSerializer.Serialize(requestBody),
                Encoding.UTF8,
                "application/json");

            var response = await _httpClient.PostAsync(url, content);
            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error marking message as read: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Gets message templates for the business account
    /// </summary>
    public async Task<string> GetMessageTemplatesAsync()
    {
        try
        {
            var url = $"{BaseUrl}/{_businessAccountId}/message_templates?access_token={_accessToken}";
            var response = await _httpClient.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting message templates: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Creates a new message template
    /// </summary>
    public async Task<bool> CreateMessageTemplateAsync(
        string templateName,
        string templateText,
        string category = "TRANSACTIONAL",
        string language = "pt_BR")
    {
        try
        {
            var requestBody = new
            {
                name = templateName,
                language = language,
                category = category,
                components = new[]
                {
                    new
                    {
                        type = "BODY",
                        text = templateText
                    }
                }
            };

            var url = $"{BaseUrl}/{_businessAccountId}/message_templates?access_token={_accessToken}";
            var content = new StringContent(
                JsonSerializer.Serialize(requestBody),
                Encoding.UTF8,
                "application/json");

            var response = await _httpClient.PostAsync(url, content);
            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating message template: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Gets the phone number details
    /// </summary>
    public async Task<string> GetPhoneNumberDetailsAsync()
    {
        try
        {
            var url = $"{BaseUrl}/{_phoneNumberId}?access_token={_accessToken}";
            var response = await _httpClient.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting phone number details: {ex.Message}");
            return null;
        }
    }

    private async Task<string?> SendMessageAsync(object requestBody)
    {
        try
        {
            var url = $"{BaseUrl}/{_phoneNumberId}/messages?access_token={_accessToken}";
            var jsonContent = JsonSerializer.Serialize(requestBody);
            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(url, content);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"WhatsApp API Error: {response.StatusCode} - {responseContent}");
                return null;
            }

            var result = JsonSerializer.Deserialize<WhatsAppResponse>(responseContent);
            var wamid = result?.Messages?.FirstOrDefault()?.Id;
            return wamid;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending message: {ex.Message}");
            return null;
        }
    }

    private object[] ConvertParametersToList(string[] parameters)
    {
        var result = new object[parameters.Length];
        for (int i = 0; i < parameters.Length; i++)
        {
            result[i] = new { type = "text", text = parameters[i] };
        }
        return result;
    }

    private class WhatsappConfig
    {
        [JsonPropertyName("whatsapp")]
        public WhatsappCredentials Whatsapp { get; set; }
    }

    private class WhatsappCredentials
    {
        [JsonPropertyName("accessToken")]
        public string AccessToken { get; set; }

        [JsonPropertyName("phoneNumberId")]
        public string PhoneNumberId { get; set; }
    }

    private class WhatsAppResponse
    {
        [JsonPropertyName("messages")]
        public List<WhatsAppResponseMessage>? Messages { get; set; }
    }

    private class WhatsAppResponseMessage
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;
    }

}
