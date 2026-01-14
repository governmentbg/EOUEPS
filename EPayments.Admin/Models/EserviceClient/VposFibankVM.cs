using EPayments.Common.DataObjects;
using System.ComponentModel.DataAnnotations;
using System.Web;

namespace EPayments.Admin.Models.EserviceClient
{
    public class VposFibankVM
    {
        public FormMode Mode { get; set; }

        public int EserviceClientId { get; set; }

        public string AisName { get; set; }
        public string DepartmentName { get; set; }

        public HttpPostedFileBase FiBankVposAccessKeystoreFile { get; set; }
        public string FiBankVposAccessKeystoreFilename { get; set; }

        [Required(ErrorMessage = "Полето „Ключ за достъп на ПИБ клиент“ е задължително.")]
        public string FiBankVposAccessKeystorePassword { get; set; }

        public AuthRequestDO TestVposRequestDO { get; set; }
        public string TestVposPostUrl { get; set; }

        public VposFibankVM()
        {
        }

        public VposFibankVM(FormMode mode,  Model.Models.EserviceClient eserviceClient)
        {
            this.Mode = mode;
            this.EserviceClientId = eserviceClient.EserviceClientId;

            this.AisName = eserviceClient.AisName;
            this.DepartmentName = eserviceClient.Department?.Name;

            if (mode != FormMode.Create)
            {
                this.FiBankVposAccessKeystoreFilename = eserviceClient.FiBankVposAccessKeystoreFilename;
                this.FiBankVposAccessKeystorePassword = eserviceClient.FiBankVposAccessKeystorePassword;
            }
        }
    }
}