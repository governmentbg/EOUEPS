using EPayments.Common.Helpers;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Admin.Models.EserviceClient
{
    public class EformsUserVM
    {
        public FormMode Mode { get; set; }

        public int EserviceClientId { get; set; }
        public int EserviceAdminUserId { get; set; }

        public string AisName { get; set; }
        public string DepartmentName { get; set; }

        [Required(ErrorMessage = "Полето „Потребителско име“ е задължително.")]
        [RegularExpression(@"^[-_A-Za-z0-9]*$", ErrorMessage = "Полето „Потребителско име“ може да съдържа само символи на латиница, цифри и символите „_“, „-“.")]
        public string Username { get; set; }

        [Required(ErrorMessage = "Полето „Име“ е задължително.")]
        public string Name { get; set; }

        [RequiredIf("IsPasswordRequired", ErrorMessage = "Полето „Парола за достъп“ е задължително.")]
        [StringLength(100, MinimumLength = 8, ErrorMessage = "Полето „Парола за достъп“ трябва да съдържа най-малко 8 символа.")]
        public string Password { get; set; }

        public bool IsPasswordRequired
        {
            get
            {
                return this.Mode == FormMode.Create;
            }
        }

        public EformsUserVM()
        {
        }

        public EformsUserVM(FormMode mode, Model.Models.EserviceClient eserviceClient)
        {
            this.Mode = mode;
            this.EserviceClientId = eserviceClient.EserviceClientId;

            this.AisName = eserviceClient.AisName;
            this.DepartmentName = eserviceClient.Department?.Name;
        }

        public EformsUserVM(FormMode mode, Model.Models.EserviceAdminUser eserviceAdminUser, Model.Models.EserviceClient eserviceClient)
        {
            this.Mode = mode;
            this.EserviceClientId = eserviceAdminUser.ReferringEserviceClientId.Value;
            this.EserviceAdminUserId = eserviceAdminUser.EserviceAdminUserId;

            this.AisName = eserviceClient.AisName;
            this.DepartmentName = eserviceClient.Department?.Name;

            this.Username = eserviceAdminUser.Username;
            this.Name = eserviceAdminUser.Name;
        }
    }
}