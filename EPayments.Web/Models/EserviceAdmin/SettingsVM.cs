namespace EPayments.Web.Models.EserviceAdmin
{
    public class SettingsVM
    {
        public string Email { get; set; }

        public bool InsufficientAmountNotifications { get; set; }

        public bool OverpaidAmountNotifications { get; set; }

        public bool ReferencedNotifications { get; set; }

        public bool NotReferencedNotifications { get; set; }
    }
}