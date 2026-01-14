using EPayments.Common.Helpers;
using EPayments.Model.Enums;
using System;

namespace EPayments.Data.ViewObjects.Web
{
    public class PaymentRequestObligationLog
    {
        public int PaymentRequestId { get; set; }

        public int ObligationStatusId { get; set; }

        public DateTime ChangeDate { get; set; }

        public string GetObliationStatus()
        {
            try
            {
                return Formatter.EnumToDescriptionString((ObligationStatusEnum)this.ObligationStatusId);
            }
            catch (Exception ex)
            {
                throw new InvalidCastException($"ObligationStatusId with value {ObligationStatusId} is not supported by enum ObligationStatusEnum.", ex);
            }
        }
    }
}
