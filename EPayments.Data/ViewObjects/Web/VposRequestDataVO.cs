using EPayments.Model.Enums;
using System;

namespace EPayments.Data.ViewObjects.Web
{
    public class VposRequestDataVO
    {
        public int PaymentRequestId { get; set; }
        public string PaymentRequestIdentifier { get; set; }
        public Guid PaymentRequestGid { get; set; }
        public PaymentRequestStatus PaymentRequestStatusId { get; set; }
        public decimal PaymentAmount { get; set; }
        public string PaymentReason { get; set; }
        public string VposPaymentRequestUri { get; set; }
        public int? VposClientId { get; set; }
        public Guid ЕserviceClientGid { get; set; }
        public string DskVposMerchantId { get; set; }
        public string DskVposMerchantPassword { get; set; }
        public string BoricaVposTerminalId { get; set; }
        public string BoricaVposBOReqSignCertFileName { get; set; }
        public string BoricaVposBOReqSignCertPassword { get; set; }
        public string FiBankVposAccessKeystoreFilename { get; set; }
        public string FiBankVposAccessKeystorePassword { get; set; }
    }
}
