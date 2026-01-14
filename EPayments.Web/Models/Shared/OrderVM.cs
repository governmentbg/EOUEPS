using EPayments.Common.DataObjects;
using EPayments.Common.Helpers;
using EPayments.Data.ViewObjects.Web;
using EPayments.Model.DataObjects;
using EPayments.Model.Enums;
using Newtonsoft.Json;
using System;

namespace EPayments.Web.Models.Shared
{
    public class OrderVM
    {
        public bool IsInternalAccess { get; set; }
        public AuthRequestDO ExternalRequestDO { get; set; }

        public Guid Gid { get; set; }
        public string ServiceProviderName { get; set; }
        public string BankName { get; set; }
        public string IBAN { get; set; }
        public string BIC { get; set; }
        public string PaymentAmount { get; set; }
        public string PaymentReason { get; set; }
        public string DocumentType { get; set; }
        public string DocumentNumber { get; set; }
        public DateTime CreateDate { get; set; }
        public string DocumentDate { get; set; }
        public string PaymentPeriodStartDate { get; set; }
        public string PaymentPeriodEndDate { get; set; }
        public string ApplicantName { get; set; }
        public string ApplicantBulstat { get; set; }
        public string ApplicantEgnLnch { get; set; }
        public string PaymentMethod { get; set; }
        public string VposAuthorizationId { get; set; }
        public string PaymentRequestIdentifier { get; set; }
        public string PaymentRequestStatusChangeTime { get; set; }
        public PaymentRequestStatus PaymentRequestStatusId { get; set; }

        public bool ShowDetailsForm { get; set; }
        public bool IsEserviceAdminAccess { get; set; }
        public string CanceledRequestStatusDesc { get; set; }
        public string ObligationType { get; set; }
        public string ObligationTypeCode { get; set; }
        public int ObligationTypeAlgorithmId { get; set; }
        public string AdditionalInformation { get; set; }
        public MDT_ExtendedInfoJson AdditionalInfoModel { get; set; }

        public OrderVM(PaymentOrderVO paymentOrderVO, AuthRequestDO externalRequestDO = null, bool isEserviceAdminAccess = false)
        {
            this.IsInternalAccess = externalRequestDO == null;
            this.ExternalRequestDO = externalRequestDO;
            this.IsEserviceAdminAccess = isEserviceAdminAccess;

            this.ShowDetailsForm =
                paymentOrderVO.PaymentRequestStatusId == PaymentRequestStatus.Canceled ||
                paymentOrderVO.PaymentRequestStatusId == PaymentRequestStatus.Suspended ||
                paymentOrderVO.PaymentRequestStatusId == PaymentRequestStatus.Expired;

            if (paymentOrderVO.PaymentRequestStatusId == PaymentRequestStatus.Canceled)
            {
                this.CanceledRequestStatusDesc = String.Format("Отказано от {0}", paymentOrderVO.ApplicantName ?? String.Empty);
            }
            else
            {
                this.CanceledRequestStatusDesc = Formatter.EnumToDescriptionString(paymentOrderVO.PaymentRequestStatusId);
            }

            this.Gid = paymentOrderVO.Gid;
            this.ServiceProviderName = paymentOrderVO.ServiceProviderName;
            this.BankName = paymentOrderVO.BankName;
            this.IBAN = paymentOrderVO.IBAN;
            this.BIC = paymentOrderVO.BIC;
            this.PaymentAmount = Formatter.DecimalToTwoDecimalPlacesFormat(paymentOrderVO.PaymentAmount);
            this.PaymentReason = paymentOrderVO.PaymentReason;
            this.DocumentType = paymentOrderVO.DocumentType;
            this.DocumentNumber = paymentOrderVO.DocumentNumber;
            this.CreateDate = paymentOrderVO.CreateDate;
            this.DocumentDate = Formatter.DateToBgFormatWithoutYearSuffix(paymentOrderVO.DocumentDate);
            this.PaymentPeriodStartDate = Formatter.DateToBgFormatWithoutYearSuffix(paymentOrderVO.PaymentPeriodStartDate);
            this.PaymentPeriodEndDate = Formatter.DateToBgFormatWithoutYearSuffix(paymentOrderVO.PaymentPeriodEndDate);
            this.ApplicantName = paymentOrderVO.ApplicantName;
            this.ApplicantBulstat = String.Empty;
            this.ApplicantEgnLnch = paymentOrderVO.ApplicantUin;
            this.PaymentMethod = paymentOrderVO.PaymentRequestStatusId.ToPaymentMethod(paymentOrderVO.IsVposAuthorized);
            this.VposAuthorizationId = paymentOrderVO.VposAuthorizationId;
            this.PaymentRequestIdentifier = paymentOrderVO.PaymentRequestIdentifier;
            this.PaymentRequestStatusChangeTime = Formatter.DateToBgFormatWithoutYearSuffix(paymentOrderVO.PaymentRequestStatusChangeTime);
            this.PaymentRequestStatusId = paymentOrderVO.PaymentRequestStatusId;
            this.ObligationType = paymentOrderVO.ObligationType;
            this.ObligationTypeCode = paymentOrderVO.ObligationTypeCode;
            this.ObligationTypeAlgorithmId = paymentOrderVO.ObligationTypeAlgorithmId;
            this.AdditionalInformation = paymentOrderVO.AdditionalInformation;
            if (!string.IsNullOrEmpty(this.AdditionalInformation))
            {
                try
                {
                    this.AdditionalInfoModel = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(paymentOrderVO.AdditionalInformation));
                }
                catch (Exception)
                {
                }
            }
        }
    }
}