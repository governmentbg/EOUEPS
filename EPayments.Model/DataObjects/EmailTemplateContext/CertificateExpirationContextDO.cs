using EPayments.Common.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Model.DataObjects.EmailTemplateContext
{
    public class CertificateExpirationContextDO
    {
        public string EserviceClientId { get; set; }
        public string EserviceClientAisName { get; set; }
        public string CertificateName { get; set; }
        public string CertificateValidTo { get; set; }
        public bool IsRequestSignCertificate { get; set; }

        public CertificateExpirationContextDO(int eserviceClientId, string eserviceClientAisName, string certificateName, DateTime certificateValidTo, bool isRequestSignCertificate)
        {
            this.EserviceClientId = eserviceClientId.ToString();
            this.EserviceClientAisName = eserviceClientAisName;
            this.CertificateName = certificateName;
            this.CertificateValidTo = Formatter.DateTimeToBgFormat(certificateValidTo);
            this.IsRequestSignCertificate = isRequestSignCertificate;
        }
    }
}