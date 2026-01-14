using EPayments.Web.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPayments.Web.Models.Home
{
    public class FeedbackVM
    {
        public string Name { get; set; }

        [StringLength(100, ErrorMessage = "Стойноста за „Електронна поща“ не трябва да надвишава 100 символа.")]
        [RegularExpression(RegexExpressions.Email, ErrorMessage = "Невалиден адрес на електронна поща.")]
        public string Email { get; set; }

        public string Phone { get; set; }

        [Required(ErrorMessage = "Изборът в полето „Съобщение за“ е задължителен.")]
        public string MessageType { get; set; }

        [Required(ErrorMessage = "Полето „Описание“ е задължително.")]
        [StringLength(4000, ErrorMessage = "Текстът в полето „Описание“ не може да съдържа повече от 4000 символа.")]
        public string Message { get; set; }

        public string Captcha { get; set; }
    }
}