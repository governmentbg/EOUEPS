using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Home
{
    public class AccessByCodeVM
    {
        [Required(ErrorMessage = "Полето „Код за плащане“ е задължително.")]
        public string AccessCode { get; set; }

        public string Captcha { get; set; }
    }
}