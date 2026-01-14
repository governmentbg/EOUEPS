using EPayments.Model.Models;
using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Settings
{
    public class EditVM
    {
        [StringLength(100, ErrorMessage = "Стойноста за „Електронна поща“ не трябва да надвишава 100 символа.")]
        [RegularExpression(RegexExpressions.Email, ErrorMessage = "Невалиден адрес на електронна поща.")]
        public string Email { get; set; }

        public bool RequestNotifications { get; set; }

        public bool StatusNotifications { get; set; }
        public bool StatusObligationNotifications { get; set; }

        public bool AccessCodeNotifications { get; set; }
    }
}