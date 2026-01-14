using EPayments.Common.Helpers;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Admin.Models.EserviceClient
{
    public class HierarchicalFormVM
    {
        [Required()]
        public int EserviceClinetId { get; set; }

        [DisplayName("Избери разпоредител")]
        [NotEqualToProperty(nameof(EserviceClinetId), ErrorMessage = "АИС клиента не може сам да си е разпоредител")]
        public int? TargetId { get; set; }
    }
}