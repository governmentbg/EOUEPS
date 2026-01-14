using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Data.ViewObjects.Web.APGModels.Requests
{
    public class APGWStatusCheckRequestDataVO : APGWRequestBase
    {
        private string _rrn;
        private string _intRef;

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

        [Display(Name = "TRTYPE")]
        public override TrTypeEnum TrType => TrTypeEnum.Check;

        [Display(Name = "TRAN_TRTYPE")]
        public TrTypeEnum TRAN_TRTYPE { get; set; }

        public override ICollection<KeyValuePair<string, string>> RequestFields()
        {
            List<KeyValuePair<string, string>> fields = new List<KeyValuePair<string, string>>()
            {
                new KeyValuePair<string, string>(base.GetFieldName(() => this.RRN, nameof(this.RRN)), this.RRN),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.INT_REF, nameof(this.INT_REF)), this.INT_REF),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Terminal, nameof(this.Terminal)), this.Terminal),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.TrType, nameof(this.TrType)), ((int)this.TrType).ToString()),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Order, nameof(this.Order)), this.Order),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.TRAN_TRTYPE, nameof(this.TRAN_TRTYPE)), ((int)this.TRAN_TRTYPE).ToString()),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.Nonce, nameof(this.Nonce)), this.Nonce),
                new KeyValuePair<string, string>(base.GetFieldName(() => this.P_Sign, nameof(this.P_Sign)), this.P_Sign),
            };

            return fields;
        }
    }
}
