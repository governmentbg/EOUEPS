using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Account
{
    public class EAuthVM
    {
        public string SamlRequest { get; set; }
        public string RelayState { get; set; }
        public string PostUrl { get; set; }
    }
}