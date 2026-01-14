using EPayments.Common;
using EPayments.Common.Helpers;
using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Account
{
    public class EidVM
    {
        public string StartAuthenticationURL { get; set; }
        public string OA { get; set; }
        public string MoccaURL { get; set; }
        public string TemplateURL { get; set; }
        public string Target { get; set; }

        public EidVM()
        {
            this.StartAuthenticationURL = AppSettings.EPaymentsWeb_EidStartAuthenticataionUrl;
            this.OA = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress.ToLower(), MVC.Account.Name.ToLower(), MVC.Account.ActionNames.Eid.ToLower()).ToString();
            this.MoccaURL = AppSettings.EPaymentsWeb_EidMoccaUrl;
            this.TemplateURL = AppSettings.EPaymentsWeb_EidTemplateUrl;
            this.Target = "BGeID";
        }
    }
}