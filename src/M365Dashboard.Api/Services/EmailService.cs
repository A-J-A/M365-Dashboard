using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;

namespace M365Dashboard.Api.Services;

public interface IEmailService
{
    Task SendReportEmailAsync(string fromEmail, List<string> toEmails, string subject, string body, string? attachmentName = null, byte[]? attachmentContent = null);
}

public class GraphEmailService : IEmailService
{
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _configuration;
    private readonly ILogger<GraphEmailService> _logger;

    public GraphEmailService(
        GraphServiceClient graphClient,
        IConfiguration configuration,
        ILogger<GraphEmailService> logger)
    {
        _graphClient = graphClient;
        _configuration = configuration;
        _logger = logger;
    }

    public async Task SendReportEmailAsync(
        string fromEmail, 
        List<string> toEmails, 
        string subject, 
        string body, 
        string? attachmentName = null, 
        byte[]? attachmentContent = null)
    {
        _logger.LogInformation("Sending report email from {From} to {Recipients}", fromEmail, string.Join(", ", toEmails));
        var senderEmail = fromEmail;

        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = body
            },
            ToRecipients = toEmails.Select(email => new Recipient
            {
                EmailAddress = new EmailAddress { Address = email }
            }).ToList()
        };

        // Add attachment if provided
        if (!string.IsNullOrEmpty(attachmentName) && attachmentContent != null)
        {
            message.Attachments = new List<Attachment>
            {
                new FileAttachment
                {
                    Name = attachmentName,
                    ContentType = GetContentType(attachmentName),
                    ContentBytes = attachmentContent,
                    OdataType = "#microsoft.graph.fileAttachment"
                }
            };
        }

        try
        {
            var requestBody = new SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = true
            };

            await _graphClient.Users[senderEmail].SendMail.PostAsync(requestBody);
            
            _logger.LogInformation("Report email sent successfully to {Recipients}", string.Join(", ", toEmails));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send report email to {Recipients}", string.Join(", ", toEmails));
            throw;
        }
    }

    private static string GetContentType(string fileName)
    {
        var extension = Path.GetExtension(fileName).ToLowerInvariant();
        return extension switch
        {
            ".csv" => "text/csv",
            ".json" => "application/json",
            ".pdf" => "application/pdf",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            _ => "application/octet-stream"
        };
    }
}
