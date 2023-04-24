using System.Net.Mail;
using MailKit.Net.Smtp;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph.Models;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using MimeKit;
using MimeKit.IO;
using MimeKit.Text;
using Umbraco.Cms.Core.Configuration.Models;
using Umbraco.Cms.Core.Events;
using Umbraco.Cms.Core.Mail;
using Umbraco.Cms.Core.Models.Email;
using Umbraco.Cms.Core.Notifications;
using Umbraco.Cms.Infrastructure.Extensions;
using static System.Formats.Asn1.AsnWriter;
using SecureSocketOptions = MailKit.Security.SecureSocketOptions;
using SmtpClient = MailKit.Net.Smtp.SmtpClient;
using Azure.Identity;

namespace VWA.Utils
{
    public class EmailSender :  IEmailSender
    {

        // TODO: This should encapsulate a BackgroundTaskRunner with a queue to send these emails!
        private readonly IEventAggregator _eventAggregator;
        private readonly ILogger<EmailSender> _logger;
        private readonly bool _notificationHandlerRegistered;
        private readonly IConfiguration _configuration;
        private GlobalSettings _globalSettings;
        
        
        public EmailSender(
            ILogger<EmailSender> logger,
            IOptionsMonitor<GlobalSettings> globalSettings,
            IEventAggregator eventAggregator,
            IConfiguration configuration)
            : this(logger, globalSettings, eventAggregator, null, null, configuration)
        {
        }

        public EmailSender(
            ILogger<EmailSender> logger,
            IOptionsMonitor<GlobalSettings> globalSettings,
            IEventAggregator eventAggregator,
            INotificationHandler<SendEmailNotification> handler1,
            INotificationAsyncHandler<SendEmailNotification> handler2,
            IConfiguration configuration)
        {
            _logger = logger;
            _eventAggregator = eventAggregator;
            _globalSettings = globalSettings.CurrentValue;
            _notificationHandlerRegistered = handler1 is not null || handler2 is not null;
            _configuration = configuration;
            globalSettings.OnChange(x => _globalSettings = x);
        }

        /// <summary>
        ///     Sends the message async
        /// </summary>
        /// <returns></returns>
        public async Task SendAsync(EmailMessage message, string emailType) =>
            await SendAsyncInternal(message, emailType, false);

        public async Task SendAsync(EmailMessage message, string emailType, bool enableNotification) =>
            await SendAsyncInternal(message, emailType, enableNotification);

        /// <summary>
        ///     Returns true if the application should be able to send a required application email
        /// </summary>
        /// <remarks>
        ///     We assume this is possible if either an event handler is registered or an smtp server is configured
        ///     or a pickup directory location is configured
        /// </remarks>
        public bool CanSendRequiredEmail() => _globalSettings.IsSmtpServerConfigured
                                              || _globalSettings.IsPickupDirectoryLocationConfigured
                                              || _notificationHandlerRegistered
                                               || _configuration.GetSection("Umbraco")?.GetSection("CMS")?.GetSection("Global")?.GetSection("Graph")?.GetValue<string>("TenantId") != null;

        private async Task SendAsyncInternal(EmailMessage message, string emailType, bool enableNotification)
        {
            if (enableNotification)
            {
                var notification =
                    new SendEmailNotification(message.ToNotificationEmail(_globalSettings.Smtp?.From), emailType);
                await _eventAggregator.PublishAsync(notification);

                // if a handler handled sending the email then don't continue.
                if (notification.IsHandled)
                {
                    _logger.LogDebug(
                        "The email sending for {Subject} was handled by a notification handler",
                        notification.Message.Subject);
                    return;
                }
            }

            var section = _configuration.GetSection("Umbraco")?.GetSection("CMS")?.GetSection("Global")?.GetSection("Graph");

            if (section == null || string.IsNullOrEmpty(section.GetValue<string>("TenantId")))
            {
                await SendAsyncInternalSmtp(message);
            }
            else
            {
                await SendAsyncInternalGraph(message, section);
            }
        }

        private async Task SendAsyncInternalGraph(EmailMessage message, IConfigurationSection section)
        {

            string clientId = section.GetValue<string>("ClientId");
            string clientSecret = section.GetValue<string>("ClientSecret");
            string tenantId = section.GetValue<string>("TenantId");
            string objectId = section.GetValue<string>("ObjectId");
            GraphServiceClient graphClient = null;

            // Get access token for Microsoft Graph
            var app = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .Build();

            
            // Define your credentials based on the created app and user details.
            // Specify the options. In most cases we're running the Azure Public Cloud.
            var credentials = new ClientSecretCredential(
                tenantId,
                clientId,
                clientSecret,
                new TokenCredentialOptions { AuthorityHost = AzureAuthorityHosts.AzurePublicCloud });

            // Initialize Microsoft Graph client
            graphClient = new GraphServiceClient(credentials);

            // Convert Umbraco message to Microsoft Graph Message
            var graphMessage = new Message
            {
                Subject = message.Subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = message.Body
                },

                ToRecipients = message.To.Select(email =>
                {
                    return CreateRecipient(email);
                }
                ).ToList(),
            };

            if (message.From != null)
            {
                graphMessage.From = CreateRecipient(message.From);
            }

            if (message.Cc != null && message.Cc.Length > 0)
            {
                graphMessage.CcRecipients = message.Cc.Select(a => CreateRecipient(a)).ToList();
            }

            if (message.Bcc != null && message.Bcc.Length > 0)
            {
                graphMessage.BccRecipients = message.Bcc.Select(a => CreateRecipient(a)).ToList();
            }
            if (message.ReplyTo != null && message.ReplyTo.Length > 0)
            {
                graphMessage.ReplyTo = message.ReplyTo.Select(a => CreateRecipient(a)).ToList();
            }
            
            graphMessage.Attachments = new List<Microsoft.Graph.Models.Attachment>();
            foreach (EmailMessageAttachment attachment in message.Attachments!)
            {
                byte[] bytes;
                using (var memoryStream = new MemoryStream())
                {
                    attachment.Stream.CopyTo(memoryStream);
                    bytes = memoryStream.ToArray();
                }

                string base64 = Convert.ToBase64String(bytes);

                graphMessage.Attachments.Add(new Microsoft.Graph.Models.Attachment()
                {
                    OdataType = "#microsoft.graph.fileAttachment",
                    Name = attachment.FileName,
                    ContentType = "text/plain",
                    AdditionalData = new Dictionary<string, object>
                    {
                        {
                            "contentBytes" , base64
                        }
                    },
                });
            }

            try
            {             
                // Send email using Microsoft Graph
                await graphClient.Users[objectId].SendMail.PostAsync(new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody { Message = graphMessage });

            }
            catch (Exception ex)
            {              
                throw;
            }
        }

        /// <summary>
        /// Returns a Graph Recipient model with the parsed email
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        private Recipient CreateRecipient(string email)
        {
            var recipient = new Recipient()
            {
                EmailAddress = new EmailAddress { Address = email }
            };

            // email could be in the form of Name <email@host.com>
            // try to parse it.
            if (MailboxAddress.TryParse(email, out InternetAddress internetAddress))
            {
                recipient.EmailAddress.Name = internetAddress.Name;
                if (internetAddress is MailboxAddress)
                {
                    recipient.EmailAddress.Address = ((MailboxAddress)internetAddress).Address;
                }
            }
            return recipient;
        }

        private async Task SendAsyncInternalSmtp(EmailMessage message)
        {
          
            if (!_globalSettings.IsSmtpServerConfigured && !_globalSettings.IsPickupDirectoryLocationConfigured)
            {
                _logger.LogDebug(
                    "Could not send email for {Subject}. It was not handled by a notification handler and there is no SMTP configured.",
                    message.Subject);
                return;
            }

            if (_globalSettings.IsPickupDirectoryLocationConfigured &&
                !string.IsNullOrWhiteSpace(_globalSettings.Smtp?.From))
            {
                // The following code snippet is the recommended way to handle PickupDirectoryLocation.
                // See more https://github.com/jstedfast/MailKit/blob/master/FAQ.md#q-how-can-i-send-email-to-a-specifiedpickupdirectory
                do
                {
                    var path = Path.Combine(_globalSettings.Smtp.PickupDirectoryLocation!, Guid.NewGuid() + ".eml");
                    Stream stream;

                    try
                    {
                        stream = File.Open(path, FileMode.CreateNew);
                    }
                    catch (IOException)
                    {
                        if (File.Exists(path))
                        {
                            continue;
                        }

                        throw;
                    }

                    try
                    {
                        using (stream)
                        {
                            using var filtered = new FilteredStream(stream);
                            filtered.Add(new SmtpDataFilter());

                            FormatOptions options = FormatOptions.Default.Clone();
                            options.NewLineFormat = NewLineFormat.Dos;

                            await message.ToMimeMessage(_globalSettings.Smtp.From).WriteToAsync(options, filtered);
                            filtered.Flush();
                            return;
                        }
                    }
                    catch
                    {
                        File.Delete(path);
                        throw;
                    }
                }
                while (true);
            }

            using var client = new MailKit.Net.Smtp.SmtpClient();

            await client.ConnectAsync(
                _globalSettings.Smtp!.Host,
                _globalSettings.Smtp.Port,
                (SecureSocketOptions)(int)_globalSettings.Smtp.SecureSocketOptions);

            if (!string.IsNullOrWhiteSpace(_globalSettings.Smtp.Username) &&
                !string.IsNullOrWhiteSpace(_globalSettings.Smtp.Password))
            {
                await client.AuthenticateAsync(_globalSettings.Smtp.Username, _globalSettings.Smtp.Password);
            }

            var mailMessage = message.ToMimeMessage(_globalSettings.Smtp.From);

            if (_globalSettings.Smtp.DeliveryMethod == SmtpDeliveryMethod.Network)
            {
                await client.SendAsync(mailMessage);
            }
            else
            {
                client.Send(mailMessage);
            }

            await client.DisconnectAsync(true);
        }
    }


    internal static class EmailMessageExtensions
    {
        public static MimeMessage ToMimeMessage(this EmailMessage mailMessage, string configuredFromAddress)
        {
            var fromEmail = string.IsNullOrEmpty(mailMessage.From) ? configuredFromAddress : mailMessage.From;

            if (!InternetAddress.TryParse(fromEmail, out InternetAddress fromAddress))
            {
                throw new ArgumentException(
                    $"Email could not be sent.  Could not parse from address {fromEmail} as a valid email address.");
            }

            var messageToSend = new MimeMessage { From = { fromAddress }, Subject = mailMessage.Subject };

            AddAddresses(messageToSend, mailMessage.To, x => x.To, true);
            AddAddresses(messageToSend, mailMessage.Cc, x => x.Cc);
            AddAddresses(messageToSend, mailMessage.Bcc, x => x.Bcc);
            AddAddresses(messageToSend, mailMessage.ReplyTo, x => x.ReplyTo);

            if (mailMessage.HasAttachments)
            {
                var builder = new BodyBuilder();
                if (mailMessage.IsBodyHtml)
                {
                    builder.HtmlBody = mailMessage.Body;
                }
                else
                {
                    builder.TextBody = mailMessage.Body;
                }

                foreach (EmailMessageAttachment attachment in mailMessage.Attachments!)
                {
                    builder.Attachments.Add(attachment.FileName, attachment.Stream);
                }

                messageToSend.Body = builder.ToMessageBody();
            }
            else
            {
                messageToSend.Body =
                    new TextPart(mailMessage.IsBodyHtml ? TextFormat.Html : TextFormat.Plain) { Text = mailMessage.Body };
            }

            return messageToSend;
        }

        public static NotificationEmailModel ToNotificationEmail(
            this EmailMessage emailMessage,
            string? configuredFromAddress)
        {
            var fromEmail = string.IsNullOrEmpty(emailMessage.From) ? configuredFromAddress : emailMessage.From;

            NotificationEmailAddress? from = ToNotificationAddress(fromEmail);

            return new NotificationEmailModel(
                from,
                GetNotificationAddresses(emailMessage.To),
                GetNotificationAddresses(emailMessage.Cc),
                GetNotificationAddresses(emailMessage.Bcc),
                GetNotificationAddresses(emailMessage.ReplyTo),
                emailMessage.Subject,
                emailMessage.Body,
                emailMessage.Attachments,
                emailMessage.IsBodyHtml);
        }

        private static void AddAddresses(MimeMessage message, string?[]? addresses, Func<MimeMessage, InternetAddressList> addressListGetter, bool throwIfNoneValid = false)
        {
            var foundValid = false;
            if (addresses != null)
            {
                foreach (var address in addresses)
                {
                    if (InternetAddress.TryParse(address, out InternetAddress internetAddress))
                    {
                        addressListGetter(message).Add(internetAddress);
                        foundValid = true;
                    }
                }
            }

            if (throwIfNoneValid && foundValid == false)
            {
                throw new InvalidOperationException("Email could not be sent. Could not parse a valid recipient address.");
            }
        }

        private static NotificationEmailAddress ToNotificationAddress(string? address)
        {
            if (InternetAddress.TryParse(address, out InternetAddress internetAddress))
            {
                if (internetAddress is MailboxAddress mailboxAddress)
                {
                    return new NotificationEmailAddress(mailboxAddress.Address, internetAddress.Name);
                }
            }

            return null;
        }

        private static IEnumerable<NotificationEmailAddress> GetNotificationAddresses(IEnumerable<string?>? addresses)
        {
            if (addresses is null)
            {
                return null;
            }

            var notificationAddresses = new List<NotificationEmailAddress>();

            foreach (var address in addresses)
            {
                NotificationEmailAddress? notificationAddress = ToNotificationAddress(address);
                if (notificationAddress is not null)
                {
                    notificationAddresses.Add(notificationAddress);
                }
            }

            return notificationAddresses;
        }
    }

}
