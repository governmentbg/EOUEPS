using System;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Data.ViewObjects.Web.APGModels
{
    public abstract class APGWBaseData
    {
        private string _order;
        private string _terminal;

        protected string DateForm => "yyyyMMddHHmmss";
        protected string GMTFormat => "zz";

        [Display(Name = "TRTYPE")]
        public virtual TrTypeEnum TrType { get; set; }

        [Display(Name = "ORDER")]
        public string Order 
        { 
            get { return this._order; }
            set
            {
                if (!this.ValidateRequiredString(value, 6))
                {
                    throw new ArgumentException("Order value is required, and it maximum length is 6.");
                }

                this._order = value;
            }
        }
        
        [Display(Name = "TERMINAL")]
        public string Terminal 
        {
            get { return this._terminal; }
            set
            {
                if (!this.ValidateRequiredString(value, 8))
                {
                    throw new ArgumentException("Terminal value is required, and it maximum length is 8.");
                }

                this._terminal = value;
            }
        }

        [Display(Name = "P_SIGN")]
        public string P_Sign { get; set; }

        [Display(Name = "NONCE")]
        public virtual string Nonce { get; set;  }

        public DateTime Date { get; set; }

        [Display(Name = "TIMESTAMP")]
        public virtual string TimeStamp
        {
            get { return Date.ToString(DateForm); }
            set { this.Date = DateTime.ParseExact(value, DateForm, System.Globalization.CultureInfo.InvariantCulture); }
        }

        [Display(Name = "MERCH_GMT")]
        public string MerchantGMT => Date.ToLocalTime().ToString(GMTFormat);

        protected string AddNextKeyToPSign(string nextKey)
        {
            if (nextKey == "-") return nextKey;
            return !string.IsNullOrWhiteSpace(nextKey) ?
                nextKey.Length.ToString() + nextKey:
                "-";
        }

        protected bool ValidateString(string value, int maxLength)
        {
            if (value == null)
            {
                return true;
            }

            if (value.Length > maxLength)
            {
                return false;
            }

            return true;
        }

        protected bool ValidateRequiredString(string value, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (value.Length > maxLength)
            {
                return false;
            }

            return true;
        }
    }
}
