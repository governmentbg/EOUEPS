using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Data.ViewObjects.Web.APGModels.Requests
{
    public class APGWReversalAuthorizationRequestDataVO : APGWRequestBase
    {
        private string _rrn;
        private string _intRef;
        private string _merchant;
        private string _description;
        private string _merchanName;
        private string _merchantUrl;
        private string _email;
        private string _adCustomBorOrderId;
        private string _addendum;

        public string RRN
        {
            get { return _rrn; }
            set
            {
                if (!base.ValidateString(value, 12))
                {
                    throw new ArgumentException("RRN maximum length is 12.");
                }

                _rrn = value;
            }
        }

        public string INT_REF
        {
            get { return _intRef; }
            set
            {
                if (!base.ValidateRequiredString(value, 32))
                {
                    throw new ArgumentException("INT_REF value is required, and it maximum length is 32.");
                }

                _intRef = value;
            }
        }

        [Display(Name = "DESC")]
        public string DESC
        {
            get { return this._description; }
            set
            {
                if (!this.ValidateString(value, 50))
                {
                    throw new ArgumentException("Description maximum length is 50.");
                }

                this._description = value;
            }
        }

        [Display(Name = "MERCHANT_NAME")]
        public string Merchan_Name
        {
            get { return this._merchanName; }
            set
            {
                if (!this.ValidateString(value, 80))
                {
                    throw new ArgumentException("Merchan name maximum length is 80.");
                }

                this._merchanName = value;
            }
        }

        [Display(Name = "MERCH_URL")]
        public string Merchant_Url
        {
            get { return this._merchantUrl; }
            set
            {
                if (!this.ValidateString(value, 250))
                {
                    throw new ArgumentException("Merchan url maximum length is 250.");
                }

                this._merchantUrl = value;
            }
        }

        [Display(Name = "MERCHANT")]
        public string Merchant
        {
            get { return this._merchant; }
            set
            {
                if (!this.ValidateString(value, 15))
                {
                    throw new ArgumentException("Merchan maximum length is 15.");
                }

                if (value.Length < 10)
                {
                    throw new ArgumentException("Merchan minimum length is 10.");
                }

                this._merchant = value;
            }
        }

        [Display(Name = "EMAIL")]
        public string Email
        {
            get { return this._email; }
            set
            {
                if (!this.ValidateString(value, 80))
                {
                    throw new ArgumentException("Email maximum length is 80.");
                }

                this._email = value;
            }
        }

        [Display(Name = "TRTYPE")]
        public override TrTypeEnum TrType => TrTypeEnum.AuthorizationReversal;

        [Display(Name = "AD.CUST_BOR_ORDER_ID")]
        public string AdCustomBorOrderId
        {
            get { return this._adCustomBorOrderId; }
            set
            {
                if (!this.ValidateString(value, 22))
                {
                    throw new ArgumentException("AdCustomBorOrderId maximum length is 22.");
                }

                this._adCustomBorOrderId = value;
            }
        }

        [Display(Name = "COUNTRY")]
        public string Country => "BG";

        [Display(Name = "ADDENDUM")]
        public string ADDENDUM
        {
            get { return _addendum; }
            set
            {
                if (!base.ValidateString(value, 5))
                {
                    throw new ArgumentException("ADDENDUM maximum length is 5.");
                }

                this._addendum = value;
            }
        }

        public override ICollection<KeyValuePair<string, string>> RequestFields()
        {
            List<KeyValuePair<string, string>> fields = new List<KeyValuePair<string, string>>()
            {
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Amount, nameof(this.Amount)), EPayments.Common.Helpers.Formatter.DecimalToTwoDecimalPlacesFormatNoSpaces(this.Amount)),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Currency, nameof(this.Currency)), this.Currency),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.RRN, nameof(this.RRN)), this.RRN),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.INT_REF, nameof(this.INT_REF)), this.INT_REF),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.DESC, nameof(this.DESC)), this.DESC),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Terminal, nameof(this.Terminal)), this.Terminal),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Merchan_Name, nameof(this.Merchan_Name)), this.Merchan_Name),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Merchant_Url, nameof(this.Merchant_Url)), this.Merchant_Url),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Merchant, nameof(this.Merchant)), this.Merchant),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Email, nameof(this.Email)), this.Email),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.TrType, nameof(this.TrType)), ((int)this.TrType).ToString()),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Order, nameof(this.Order)), this.Order),
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
