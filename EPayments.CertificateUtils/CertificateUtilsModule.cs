using Autofac;
using EPayments.CertificateUtils.Extractions;
using EPayments.CertificateUtils.Extractions.IssuerExtractions;
using EPayments.CertificateUtils.Extractor;

namespace EPayments.CertificateUtils
{
    public class CertificateUtilsModule : Module
    {
        protected override void Load(ContainerBuilder moduleBuilder)
        {
            moduleBuilder.RegisterType<CertificateExtractor>().As<ICertificateExtractor>();

            //CertificateDO Extractions
            moduleBuilder.RegisterType<BTrustCertificateExtraction>().As<ICertificateExtraction>();
            moduleBuilder.RegisterType<InfoNotaryCertificateExtraction>().As<ICertificateExtraction>();
            moduleBuilder.RegisterType<SepCertificateExtraction>().As<ICertificateExtraction>();
            moduleBuilder.RegisterType<SpektarCertificateExtraction>().As<ICertificateExtraction>();
            moduleBuilder.RegisterType<StampITCertificateExtraction>().As<ICertificateExtraction>();
            moduleBuilder.RegisterType<EvrotrustCertificateExtraction>().As<ICertificateExtraction>();
        }
    }
}
