using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace EPayments.Data.ViewObjects.Web.APGModels.Requests
{
    public class APGWSoftRequestDataVO : APGWPaymentRequestDataVO
    {
        private string _description;

        [Display(Name = "M_INFO")]
        public virtual string M_Info => Convert.ToBase64String(Encoding.UTF8.GetBytes(string.Format("{{ \"{0}\":\"{1}\" }}", "threeDSRequestorChallengeInd", "04")));

        [Display(Name = "TRTYPE")]
        public override TrTypeEnum TrType => TrTypeEnum.Payment;

        public virtual string DSRequestorChallengeIndName { get; set; } = "threeDSRequestorChallengeInd";

        public virtual string DSRequestorChallengeIndCode { get; set; } = "04";

        public override ICollection<KeyValuePair<string, string>> RequestFields()
        {
            List<KeyValuePair<string, string>> fields = new List<KeyValuePair<string, string>>()
            {
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Amount, nameof(this.Amount)), EPayments.Common.Helpers.Formatter.DecimalToTwoDecimalPlacesFormatNoSpaces(this.Amount)),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Currency, nameof(this.Currency)), this.Currency),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.DESC, nameof(this.DESC)), this.DESC),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Terminal, nameof(this.Terminal)), this.Terminal),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Merchant_Name, nameof(this.Merchant_Name)), this.Merchant_Name),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Merchant_Url, nameof(this.Merchant_Url)), this.Merchant_Url),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Merchant, nameof(this.Merchant)), this.Merchant),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Email, nameof(this.Email)), this.Email),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.TrType, nameof(this.TrType)), ((int)this.TrType).ToString()),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Order, nameof(this.Order)), this.Order),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.M_Info, nameof(this.M_Info)), this.M_Info),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.AdCustomBorOrderId, nameof(this.AdCustomBorOrderId)), this.AdCustomBorOrderId),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Country, nameof(this.Country)), this.Country),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.TimeStamp, nameof(this.TimeStamp)), this.TimeStamp),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.MerchantGMT, nameof(this.MerchantGMT)), this.MerchantGMT),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Nonce, nameof(this.Nonce)), this.Nonce),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.ADDENDUM, nameof(this.ADDENDUM)), this.ADDENDUM),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.P_Sign, nameof(this.P_Sign)), this.P_Sign),
            };

            return fields;
        }
    }
}
