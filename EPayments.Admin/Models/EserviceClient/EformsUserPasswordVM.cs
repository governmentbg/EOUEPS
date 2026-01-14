using System.ComponentModel.DataAnnotations;

namespace EPayments.Admin.Models.EserviceClient
{
    public class EformsUserPasswordVM
    {
        public int EserviceClientId { get; set; }
        public int EserviceAdminUserId { get; set; }

        public string AisName { get; set; }
        public string DepartmentName { get; set; }

        public string Username { get; set; }
        public string Name { get; set; }

        [Required(ErrorMessage = "Полето „Парола за достъп“ е задължително.")]
        [StringLength(100, MinimumLength = 8, ErrorMessage = "Полето „Парола за достъп“ трябва да съдържа най-малко 8 символа.")]
        public string Password { get; set; }

        public EformsUserPasswordVM()
        {
        }

        public EformsUserPasswordVM(Model.Models.EserviceAdminUser eserviceAdminUser, Model.Models.EserviceClient eserviceClient)
        {
            this.EserviceClientId = eserviceAdminUser.ReferringEserviceClientId.Value;
            this.EserviceAdminUserId = eserviceAdminUser.EserviceAdminUserId;

            this.AisName = eserviceClient.AisName;
            this.DepartmentName = eserviceClient.Department?.Name;

            this.Username = eserviceAdminUser.Username;
            this.Name = eserviceAdminUser.Name;
        }
    }
}