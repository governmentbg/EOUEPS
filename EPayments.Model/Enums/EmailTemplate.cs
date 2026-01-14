using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Model.Enums
{
    public enum EmailTemplate
    {
        FeedbackMessage = 1,
        NewPaymentRequestMessage = 2,
        StatusChangedPaymentRequestMessage = 3,
        CertificateExpirationMessage = 4,
        SharePaymentMessage = 5,
        AccessCodeActivatedMessage = 6,
        StatusChangedObligationMessage = 7,
        AccessCodeApplicantActivatedMessage = 8
    }
}
