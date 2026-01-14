using EPayments.Web.Common;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Web.Models.EserviceAdmin
{
    public class EditSettingsVM
    {
        [StringLength(100, ErrorMessage = "Стойноста за „Електронна поща“ не трябва да надвишава 100 символа.")]
        [RegularExpression(RegexExpressions.Email, ErrorMessage = "Невалиден адрес на електронна поща.")]
        public string Email { get; set; }

        public bool InsufficientAmountNotifications { get; set; }

        public bool OverpaidAmountNotifications { get; set; }

        public bool ReferencedNotifications { get; set; }

        public bool NotReferencedNotifications { get; set; }
    }
}