using EPayments.Data.ViewObjects.Web;
using EPayments.Web.Common;
using EPayments.Web.DataObjects;
using EPayments.Web.Models.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Web.Models.Payments
{
    public class ShareVM
    {
        public Guid Gid { get; set; }

        public string AccessCode { get; set; }

        public string Link { get; set; }

        [Required(ErrorMessage = "Полето „Електронна поща“ е задължително.")]
        [StringLength(100, ErrorMessage = "Стойноста за „Електронна поща“ не трябва да надвишава 100 символа.")]
        [RegularExpression(RegexExpressions.Email, ErrorMessage = "Невалиден адрес на електронна поща.")]
        public string Email { get; set; }
    }
}