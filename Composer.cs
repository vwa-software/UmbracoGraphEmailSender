using www.Services;
using www.Smidge;
using Smidge.FileProcessors;
using Umbraco.Cms.Core.Composing;
using Umbraco.Cms.Core.Notifications;
using Umbraco.Cms.Infrastructure.Examine;
using Umbraco.Forms.Core.Providers;
using UmbracoFormsExtensions;
using www.Controllers;
using Umbraco.Cms.Core.Mail;

namespace VWA.Infrastructure
{
    public class Composer : IComposer
    {
        /// <summary>Compose.</summary>
        public void Compose(IUmbracoBuilder builder)
        {
        
            // use our own email sender
            builder.Services.Remove(new ServiceDescriptor(typeof(IEmailSender), typeof(Umbraco.Cms.Infrastructure.Mail.EmailSender), ServiceLifetime.Singleton));
            builder.Services.AddSingleton<IEmailSender, ETC.HKliving.Shared.EmailSender>();

        }
    }
}
