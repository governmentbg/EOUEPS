using System.ComponentModel.DataAnnotations;

namespace EPayments.Web.Models.Home
{
    public class AccessByEserviceAdminVM
    {
        [Required(ErrorMessage = "Полето „Потребител“ е задължително.")]
        public string Username { get; set; }

        [Required(ErrorMessage = "Полето „Парола“ е задължително.")]
        public string Password { get; set; }

        public string Captcha { get; set; }
    }
}