using EPayments.Common;
using EPayments.Common.Helpers;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq.Expressions;
using System.Reflection;

namespace EPayments.Data.ViewObjects.Web.APGModels
{
    public abstract class APGWRequestBase : APGWBaseData
    {
        private string _nonce = null;

        public virtual string PostUrl
        {
            get { return AppSettings.EPaymentsWeb_CentralVposUrl; }
        }

        [Display(Name = "AMOUNT")]
        public decimal Amount { get; set; }

        [Display(Name = "CURRENCY")]
        public string Currency => CurrencyHelper.GetCurrencyCode();

        [Display(Name = "NONCE")]
        public override string Nonce
        {
            get
            {
                if (_nonce != null)
                {
                    return _nonce;
                }

                Random random = new Random();
                _nonce = string.Empty;

                for (int i = 0; i < 16; i++)
                {
                    _nonce = _nonce + ((byte)random.Next(0, byte.MaxValue)).ToString("x2").ToUpper();
                }

                return _nonce;
            }
        }

        protected string GetFieldName<T>(Expression<Func<T>> expression, string defaultName)
        {
            MemberExpression memberExpression = expression.Body as MemberExpression;

            if (memberExpression != null)
            {
                MemberInfo memberInfo = memberExpression.Member;

                DisplayAttribute customAttribute = memberInfo.GetCustomAttribute(typeof(DisplayAttribute)) as DisplayAttribute;

                if (customAttribute != null)
                {
                    return customAttribute.Name;
                }
            }

            return defaultName;
        }

        public string GetPSignData()
        {
            string pSign = "";
            if (TrType == TrTypeEnum.Check)
            {
                pSign = base.AddNextKeyToPSign(this.Terminal) +
                base.AddNextKeyToPSign(((int)this.TrType).ToString()) +
                base.AddNextKeyToPSign(this.Order) +
                base.AddNextKeyToPSign(this.Nonce);
            }
            else
            {
                pSign = base.AddNextKeyToPSign(this.Terminal) +
                base.AddNextKeyToPSign(((int)this.TrType).ToString()) +
                base.AddNextKeyToPSign(this.Amount.ToString("#0.00", System.Globalization.CultureInfo.InvariantCulture)) +
                base.AddNextKeyToPSign(this.Currency) +
                base.AddNextKeyToPSign(this.Order) +
                base.AddNextKeyToPSign(this.TimeStamp) +
                base.AddNextKeyToPSign(this.Nonce) +
                base.AddNextKeyToPSign("-");//"-" is for the RFU (Reserved for Future use) field which is not used and implemented currently
            }

            return pSign;
        }

        public abstract ICollection<KeyValuePair<string, string>> RequestFields();
    }
}
