using EPayments.Common.DataObjects;
using System.ComponentModel.DataAnnotations;

namespace EPayments.Admin.Models.EserviceClient
{
    public class VposDskVM
    {
        public FormMode Mode { get; set; }

        public int EserviceClientId { get; set; }

        public string AisName { get; set; }
        public string DepartmentName { get; set; }

        [Required(ErrorMessage = "Полето „Идентификатор на ДСК клиент/търговец“ е задължително.")]
        public string DskVposMerchantId { get; set; }

        [Required(ErrorMessage = "Полето „Ключ на ДСК клиент/търговец“ е задължително.")]
        public string DskVposMerchantPassword { get; set; }

        public AuthRequestDO TestVposRequestDO { get; set; }
        public string TestVposPostUrl { get; set; }

        public VposDskVM()
        {
        }

        public VposDskVM(FormMode mode, Model.Models.EserviceClient eserviceClient)
        {
            this.Mode = mode;
            this.EserviceClientId = eserviceClient.EserviceClientId;

            this.AisName = eserviceClient.AisName;
            this.DepartmentName = eserviceClient.Department?.Name;

            if (mode != FormMode.Create)
            {
                this.DskVposMerchantId = eserviceClient.DskVposMerchantId;
                this.DskVposMerchantPassword = eserviceClient.DskVposMerchantPassword;
            }
        }
    }
}