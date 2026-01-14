using EPayments.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Job.Host.Email
{
    public class TemplateConfig
    {
        public string TemplateName { get; private set; }
        public string TemplateFileName { get; private set; }
        public string Sender { get; private set; }
        public string MailSubject { get; private set; }
        public bool IsBodyHtml { get; private set; }

        private static IList<TemplateConfig> AllTemplates;

        private TemplateConfig(string templateName, string templateFileName, string sender, string mailSubject, bool isBodyHtml)
        {
            this.TemplateName = templateName;
            this.TemplateFileName = templateFileName;
            this.Sender = sender;
            this.MailSubject = mailSubject;
            this.IsBodyHtml = isBodyHtml;
        }

        public static TemplateConfig Get(string name)
        {
            return AllTemplates.Single(e => e.TemplateName == name);
        }

        static TemplateConfig()
        {
            AllTemplates = new List<TemplateConfig>()
            {
                new TemplateConfig(
                    "FeedbackMessage",
                    "FeedbackMessage.cshtml",
                    AppSettings.EPaymentsJobHost_EmailJobSender,
                    "Среда за електронни плащания към доставчиците на електронни административни услуги (обратна връзка)",
                    true),

                new TemplateConfig(
                    "NewPaymentRequestMessage",
                    "NewPaymentRequestMessage.cshtml",
                    AppSettings.EPaymentsJobHost_EmailJobSender,
                    "Среда за електронни плащания към доставчиците на електронни административни услуги (ново задължение)",
                    true),

                new TemplateConfig(
                    "StatusChangedPaymentRequestMessage",
                    "StatusChangedPaymentRequestMessage.cshtml",
                    AppSettings.EPaymentsJobHost_EmailJobSender,
                    "Среда за електронни плащания към доставчиците на електронни административни услуги (промяна на статус на плащане)",
                    true),
                new TemplateConfig(
                    "StatusChangedObligationMessage",
                    "StatusChangedObligationMessage.cshtml",
                    AppSettings.EPaymentsJobHost_EmailJobSender,
                    "Среда за електронни плащания към доставчиците на електронни административни услуги (промяна на статус на задължение)",
                    true),

                new TemplateConfig(
                    "CertificateExpirationMessage",
                    "CertificateExpirationMessage.cshtml",
                    AppSettings.EPaymentsJobHost_EmailJobSender,
                    "Среда за електронни плащания към доставчиците на електронни административни услуги (изтичане на БОРИКА сертификат)",
                    true),

                new TemplateConfig(
                    "SharePaymentMessage",
                    "SharePaymentMessage.cshtml",
                    AppSettings.EPaymentsJobHost_EmailJobSender,
                    "Среда за електронни плащания към доставчиците на електронни административни услуги (Код за плащане)",
                    true),

                new TemplateConfig(
                    "AccessCodeActivatedMessage",
                    "AccessCodeActivatedMessage.cshtml",
                    AppSettings.EPaymentsJobHost_EmailJobSender,
                    "Среда за електронни плащания към доставчиците на електронни административни услуги (Код за плащане)",
                    true),

                new TemplateConfig(
                    "AccessCodeApplicantActivatedMessage",
                    "AccessCodeApplicantActivatedMessage.cshtml",
                    AppSettings.EPaymentsJobHost_EmailJobSender,
                    "Среда за електронни плащания към доставчиците на електронни административни услуги (Код за плащане)",
                    true),
            };
        }
    }
}