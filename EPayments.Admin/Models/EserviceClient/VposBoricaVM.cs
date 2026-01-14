using EPayments.Common.DataObjects;
using EPayments.Common.Helpers;
using System.ComponentModel.DataAnnotations;
using System.Web;

namespace EPayments.Admin.Models.EserviceClient
{
    public class VposBoricaVM
    {
        public FormMode Mode { get; set; }

        public int EserviceClientId { get; set; }

        public string AisName { get; set; }
        public string DepartmentName { get; set; }

        public HttpPostedFileBase BoricaVposRequestSignCertFile { get; set; }
        public string BoricaVposRequestSignCertFileName { get; set; }

        public string BoricaVposRequestSignCertPassword { get; set; }

        public string BoricaVposRequestSignCertValidTo { get; set; }

        [Required(ErrorMessage = "Полето „ID на терминал“ е задължително.")]
        public string BoricaVposTerminalId { get; set; }

        [Required(ErrorMessage = "Полето „ID на търговец“ е задължително.")]
        public string BoricaVposMerchantId { get; set; }

        public AuthRequestDO TestVposRequestDO { get; set; }
        public string TestVposPostUrl { get; set; }

        public VposBoricaVM()
        {
            
        }

        public VposBoricaVM(FormMode mode, Model.Models.EserviceClient eserviceClient)
        {
            this.Mode = mode;
            this.EserviceClientId = eserviceClient.EserviceClientId;

            this.AisName = eserviceClient.AisName;
            this.DepartmentName = eserviceClient.Department?.Name;

            if (mode != FormMode.Create)
            {
                this.BoricaVposTerminalId = eserviceClient.BoricaVposTerminalId;
                this.BoricaVposMerchantId = eserviceClient.BoricaVposMerchantId;
                this.BoricaVposRequestSignCertFileName = eserviceClient.BoricaVposRequestSignCertFileName;
                this.BoricaVposRequestSignCertPassword = eserviceClient.BoricaVposRequestSignCertPassword;
                this.BoricaVposRequestSignCertValidTo = Formatter.DateToBgFormat(eserviceClient.BoricaVposRequestSignCertValidTo.Value);
            }
        }
    }
}