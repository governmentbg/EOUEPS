using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.Common
{
    public class SessionKeys
    {
        public const string CertificateRawContent = "Session_CertificateContent";
    }

    public class TempDataKeys
    {
        public const string ErrorMessage = "TempData_ErrorMessage";
        public const string Message = "TempData_Message";
        public const string FeedbackSend = "TempData_FeedbackSend";
        public const string SharePaymentEmailSend = "TempData_SharePaymentEmailSend";
        public const string IsVposPaymentSuccessful = "TempData_IsVposPaymentSuccessful";
        public const string VposPaymentMessage = "TempData_VposPaymentMessage";
        public const string ErrorId = "TempData_ErrorId";
        public const string ErrorAttemptLogId = "TempData_ErrorAttemptLogId";
        public const string ErrorUserEgn = "TempData_ErrorUserEgn";
        public const string IsIisError = "TempData_IsIisError";
        public const string Url = "TempData_Url";
        public const string EAuthErrorMessage = "TempData_EAuthErrorMessage";
        public const string SearchPerformed = "TempData_SearchPerformed";
        public const string IsPaymentRequestError = "TempData_IsPaymentRequestError";
        public const string PaymentRequestError = "TempData_PaymentRequestError";
    }

    public class Constants
    {
        public const string CaptchaModelName = "ValidationImage";
        public const string ProcessedPaymentsFocusId = "processedPayments";
    }
}