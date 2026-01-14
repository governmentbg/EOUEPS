using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace EPayments.Common
{
    public class AppSettings
    {
        #region EPayments.Web

        public static int EPaymentsWeb_MaxSearchResultsPerPage
        {
            get
            {
                return 10;
            }
        }

        public static int EPaymentsWeb_MaxTransactionResultsPerPage
        {
            get
            {
                return 10;
            }
        }

        public static bool EPaymentsWeb_UseFakeCertificate
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:UseFakeCertificate");
            }
        }

        public static bool EPaymentsWeb_UseEAuthForLogin
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:UseEAuthForLogin");
            }
        }

        public static string EPaymentsWeb_FakeCertificateBase64Data
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:FakeCertificateBase64Data");
            }
        }

        public static bool EPaymentsWeb_NotificationUrlValidateSslCert
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:NotificationUrlValidateSslCert");
            }
        }

        public static bool EPaymentsWeb_SimulateVposPayment
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:SimulateVposPayment");
            }
        }

        public static string EPaymentsWeb_FeedbackEmail
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:FeedbackEmail");
            }
        }

        public static string EPaymentsWeb_PortalName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:PortalName");
            }
        }

        public static bool EPaymentsWeb_PaymentUseMinTestAmount
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:PaymentUseMinTestAmount");
            }
        }

        public static bool EPaymentsWeb_PaymentAllowDublicateTrackId
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:PaymentAllowDublicateTrackId");
            }
        }

        public static bool EPaymentsWeb_EnableGoogleAnalyticsLogging
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:EnableGoogleAnalyticsLogging");
            }
        }

        public static bool EPaymentsWeb_EAuthEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:EAuthEnabled");
            }
        }

        public static bool EPaymentsWeb_EAuthSkipped
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:EAuthSkipped");
            }
        }

        public static bool EPaymentsWeb_IsAuthPassEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:IsAuthPassEnabled");
            }
        }

        public static string EPaymentsWeb_EAuthRequestUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthRequestUrl");
            }
        }

        public static string EPaymentsWeb_EAuthMetadataUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthMetadataUrl");
            }
        }

        public static string EPaymentsWeb_EAuthProviderId
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthProviderId");
            }
        }

        public static string EPaymentsWeb_EAuthExtServiceId
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthExtServiceId");
            }
        }

        public static string EPaymentsWeb_EAuthExtProviderId
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthExtProviderId");
            }
        }

        public static string EPaymentsWeb_EAuthResponseSignCertificateThumbprint
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthResponseSignCertificateThumbprint");
            }
        }

        public static bool EPaymentsWeb_EAuthSkipSignatureCheck
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:EAuthSkippedSignatureCheck");
            }
        }

        public static bool EPaymentsWeb_EAuthResponseSignCertificateValidateExpirationDate
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:EAuthResponseSignCertificateValidateExpirationDate");
            }
        }

        public static string EPaymentsWeb_EAuthRequestSignCertificatePass
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthRequestSignCertificatePass");
            }
        }

        public static string EPaymentsWeb_EAuthIdpMetadata
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthIdpMetadata");
            }
        }

        public static string EPaymentsWeb_EAuthRequestSignCertificatePath
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certificates", "EAuth", GetAppConfigValue<string>("EPayments.Web:EAuthRequestSignCertificate"));
            }
        }

        public static string EPaymentsWeb_EidMoccaUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EidMoccaUrl");
            }
        }

        public static string EPaymentsWeb_EidTemplateUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EidTemplateUrl");
            }
        }

        public static string EPaymentsWeb_EidStartAuthenticataionUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EidStartAuthenticataionUrl");
            }
        }

        public static bool EPaymentsWeb_EidEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:EidEnabled");
            }
        }

        public static bool EPaymentsWeb_NoiAuthEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:NoiAuthEnabled");
            }
        }

        public static bool EPaymentsWeb_EDeliveryAuthEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:EDeliveryAuthEnabled");
            }
        }

        public static string EPaymentsWeb_EDeliveryAuthSecret
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EDeliveryAuthSecret");
            }
        }

        public static bool EPaymentsWeb_EMySpaceAuthEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:EMySpaceAuthEnabled");
            }
        }

        public static string EPaymentsWeb_EMySpaceAuthSecret
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EMySpaceAuthSecret");
            }
        }

        public static string EPaymentsWeb_EMySpaceAuthOID
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EMySpaceAuthOID");
            }
        }

        public static string EPaymentsWeb_EMySpaceAuthSecretSalt
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EMySpaceAuthSecretSalt");
            }
        }

        public static string EPaymentsWeb_NoiAuthSecretKey
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:NoiAuthSecretKey");
            }
        }

        public static string EPaymentsWeb_EAuthAdminRedirect
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EAuthAdminRedirect");
            }
        }

        public static bool EPaymentsWeb_HelpDeskSubmitEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:HelpDeskSubmitEnabled");
            }
        }

        public static string EPaymentsWeb_HelpDeskUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:HelpDeskUrl");
            }
        }

        public static string EPaymentsWeb_HelpDeskUsername
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:HelpDeskUsername");
            }
        }

        public static string EPaymentsWeb_HelpDeskPassword
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:HelpDeskPassword");
            }
        }

        public static string EPaymentsWeb_HelpDeskClientId
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:HelpDeskClientId");
            }
        }

        public static string EPaymentsWeb_HelpDeskClientSecret
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:HelpDeskClientSecret");
            }
        }

        public static string EPaymentsWeb_FiBankVposMerchantHandlerUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:FiBankVposMerchantHandlerUrl");
            }
        }

        public static string EPaymentsWeb_FiBankVposClientHandlerUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:FiBankVposClientHandlerUrl");
            }
        }

        public static string EPaymentsWeb_FiBankVposTestIp
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:FiBankVposTestIp");
            }
        }

        public static string EPaymentsWeb_EpayVposKin
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EpayVposKin");
            }
        }

        public static string EPaymentsWeb_EpayVposSecret
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EpayVposSecret");
            }
        }

        public static string EPaymentsWeb_EpayVposUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EpayVposUrl");
            }
        }

        public static string EPaymentsWeb_EpayOK
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EpayOK");
            }
        }

        public static string EPaymentsWeb_EpayCancel
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EpayCancel");
            }
        }

        //CvPos Settings
        public static string EPaymentsWeb_CentralVposPrivateKeyPassphrase
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposPrivateKeyPassphrase");
            }
        }

        public static string EPaymentsWeb_CentralVposPrivateKeyFileName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposPrivateKeyFileName");
            }
        }

        public static string EPaymentJobHost_CentralVposPrivateKeyPassphrase
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CentralVposPrivateKeyPassphrase");
            }
        }
        public static string EPaymentsJobHost_CentralVposPrivateKeyFileName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CentralVposPrivateKeyFileName");
            }
        }

        public static string EPaymentsWeb_BorikaPublicKeyFileName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:BorikaPublicKeyFileName");
            }
        }

        public static string EPaymentsWeb_CentralVposUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposUrl");
            }
        }

        public static string EPaymentsWeb_CentralVposBackref
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposBackref");
            }
        }

        public static string EPaymentsWeb_CentralVposDescription
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposDescription");
            }
        }

        public static string EPaymentsWeb_CentralVposADDENDUM
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposADDENDUM");
            }
        }

        public static string EPaymentsWeb_CentralVposPrefixHelper
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposPrefixHelper");
            }
        }

        public static string EPaymentsWeb_CentralVposDevTerminalId
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposDevTerminalId");
            }
        }

        public static string EPaymentsWeb_CentralVposMerchantId
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CentralVposMerchantId");
            }
        }

        public static string EPaymentsWeb_FiBankCertificateFolder
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certificates", "FiBankVpos");
            }
        }

        public static string EPaymentsWeb_BoricaCertificateFolder
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certificates", "BoricaVpos");
            }
        }

        public static string EPaymentsWeb_BoricaCentralVposCertificateFolder
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certificates", "BoricaCvpos");
            }
        }

        public static string EPaymentsWeb_CreatePaymentRequest
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:CreatePaymentRequest");
            }
        }

        public static string EPaymentsWeb_SuspendPaymentRequest
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:SuspendPaymentRequest");
            }
        }

        public static string EPaymentsWeb_SetStatusPaidPaymentRequest
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:SetStatusPaidPaymentRequest");
            }
        }

        public static string EPaymentsWeb_MerchantName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:MerchantName");
            }
        }

        public static string EPaymentsWeb_MerchantEmail
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:MerchantEmail");
            }
        }

        public static string EPaymentsWeb_MerchantUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:MerchantUrl");
            }
        }

        #endregion

        #region EPayments.EventRegister

        public static bool EPaymentsEventRegister_Enabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.EventRegister:Enabled");
            }
        }

        #endregion

        #region EPayments.Admin

        public static string EPaymentsAdmin_EAuthReturnUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Admin:EAuthReturnUrl");
            }
        }

        public static string EPaymentsAdmin_FiBankCertificateFolder
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Admin:FiBankCertificateFolder");
            }
        }

        public static string EPaymentsAdmin_BoricaCertificateFolder
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Admin:BoricaCertificateFolder");
            }
        }

        public static string EPaymentsAdmin_BoricaDistributionAgency
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Admin:BoricaDistributionAgency");
            }
        }

        public static string EPaymentsAdmin_BoricaDistributionBICRegex
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Admin:BoricaDistributionBICRegex");
            }
        }

        public static string EPaymentsAdmin_BoricaDistributionBICErrorMessage
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Admin:BoricaDistributionBICErrorMessage");
            }
        }

        #endregion

        #region EPayments.Job.Host

        //EmailJob

        public static bool EPaymentsJobHost_EmailJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:EmailJobEnabled");
            }
        }

        public static bool EmailJobUseTestGmailServer
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:EmailJobUseTestGmailServer");
            }
        }

        public static string EPaymentsJobHost_EmailJobSender
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EmailJobSender");
            }
        }

        public static string EPaymentsJobHost_EmailJobMailServer
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EmailJobMailServer");
            }
        }

        public static int EPaymentsJobHost_EmailJobBatchSize
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EmailJobBatchSize");
            }
        }

        public static int EPaymentsJobHost_EmailJobPeriodInSeconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EmailJobPeriodInSeconds");
            }
        }

        public static int EPaymentsJobHost_EmailJobMaxFailedAttempts
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EmailJobMaxFailedAttempts");
            }
        }

        public static double EPaymentsJobHost_EmailJobFailedAttemptTimeoutInMinutes
        {
            get
            {
                return GetAppConfigValue<double>("EPayments.Job.Host:EmailJobFailedAttemptTimeoutInMinutes");
            }
        }

        public static int EPaymentsJobHost_EmailJobParallelTasks
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EmailJobParallelTasks");
            }
        }

        //EserviceNotificationJob

        public static bool EPaymentsJobHost_EserviceNotificationJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:EserviceNotificationJobEnabled");
            }
        }

        public static bool EPaymentsJobHost_EventRegisterNotificationJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:EventRegisterNotificationJobEnabled");
            }
        }

        public static int EPaymentsJobHost_EserviceNotificationJobPeriodInSeconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EserviceNotificationJobPeriodInSeconds");
            }
        }

        public static int EPaymentsJobHost_EserviceNotificationJobParallelTasks
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EserviceNotificationJobParallelTasks");
            }
        }

        public static int EPaymentsJobHost_EserviceNotificationJobBatchSize
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EserviceNotificationJobBatchSize");
            }
        }

        //EventRegisterNotificationJob

        public static int EPaymentsJobHost_EventRegisterNotificationJobPeriodInSeconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EventRegisterNotificationJobPeriodInSeconds");
            }
        }

        public static bool EPaymentsJobHost_EventRegisterNotificationJobValidateSslCert
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:EventRegisterNotificationJobValidateSslCert");
            }
        }

        public static string EPaymentsJobHost_EventRegisterNotificationJobAdminOID
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EventRegisterNotificationJobAdminOID");
            }
        }

        public static string EPaymentsJobHost_EventRegisterNotificationJobAdminLegalName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EventRegisterNotificationJobAdminLegalName");
            }
        }

        public static string EPaymentsJobHost_EventRegisterNotificationJobServiceOID
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EventRegisterNotificationJobServiceOID");
            }
        }

        public static string EPaymentsJobHost_EventRegisterNotificationJobSPOID
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EventRegisterNotificationJobSPOID");
            }
        }

        public static string EPaymentsJobHost_EventRegisterNotificationJobSPName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EventRegisterNotificationJobSPName");
            }
        }

        public static string EPaymentsJobHost_EDeliveryNotificationJobSPName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EDeliveryNotificationJobSPName");
            }
        }

        public static string EPaymentsJobHost_EventRegisterNotificationJobInformationSystemOID
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EventRegisterNotificationJobInformationSystemOID");
            }
        }

        public static string EPaymentsJobHost_EDeliveryNotificationJobInformationSystemOID
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EDeliveryNotificationJobInformationSystemOID");
            }
        }

        //ExpiredRequestJob

        public static bool EPaymentsJobHost_ExpiredRequestJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:ExpiredRequestJobEnabled");
            }
        }

        public static int EPaymentsJobHost_ExpiredRequestJobPeriodInSeconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:ExpiredRequestJobPeriodInSeconds");
            }
        }

        public static string EPaymentsJobHost_ExpiredRequestJobFeedbackEmail
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:ExpiredRequestJobFeedbackEmail");
            }
        }

        //EPaymentsJobHost_ProcessTransactionFilesJob

        public static bool EPaymentsJobHost_ProcessTransactionFilesJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:ProcessTransactionFilesJobEnabled");
            }
        }

        public static int EPaymentsJobHost_ProcessTransactionFilesJobPeriodInSeconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:ProcessTransactionFilesJobPeriodInSeconds");
            }
        }


        //EPaymentsJobHost_CVPosTransactionJob

        public static bool EPaymentsJobHost_CVPosTransactionJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:CVPosTransactionJobEnabled");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionJobPeriodInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionJobPeriodInMinutes");
            }
        }

        public static int PaymentsJobHost_CVPosTransactionJobBatchSize
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionJobBatchSize");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionJobMaxFailedAttempts
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionJobMaxFailedAttempts");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionJobFailedAttemptTimeoutInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionJobFailedAttemptTimeoutInMinutes");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionJobTimeoutInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionJobTimeoutInMinutes");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionJobTimeoutBetweenRequestsInMilliseconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionJobTimeoutBetweenRequestsInMilliseconds");
            }
        }

        public static bool EPaymentsJobHost_CVPosTransactionJobTestMode
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:CVPosTransactionJobTestMode");
            }
        }

        public static string EPaymentsJobHost_CVPosTransactionJobAgency
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CVPosTransactionJobAgency");
            }
        }

        public static string EPaymentsJobHost_CVPosTransactionJobEvent
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CVPosTransactionJobEvent");
            }
        }

        public static string CVPosTransactionJobStartTime
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CVPosTransactionJobStartTime");
            }
        }

        //EPaymentsJobHost_CVPosTransactionFixJob

        public static bool EPaymentsJobHost_CVPosTransactionFixJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:CVPosTransactionFixJobEnabled");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionFixJobPeriodInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionFixJobPeriodInMinutes");
            }
        }

        public static int PaymentsJobHost_CVPosTransactionFixJobBatchSize
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionFixJobBatchSize");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionFixJobMaxFailedAttempts
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionFixJobMaxFailedAttempts");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionFixJobFailedAttemptTimeoutInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionFixJobFailedAttemptTimeoutInMinutes");
            }
        }

        public static int EPaymentsJobHost_CVPosTransactionFixJobTimeoutBetweenRequestsInMilliseconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:CVPosTransactionFixJobTimeoutBetweenRequestsInMilliseconds");
            }
        }

        public static bool EPaymentsJobHost_CVPosTransactionFixJobTestMode
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:CVPosTransactionFixJobTestMode");
            }
        }

        public static string EPaymentsJobHost_CVPosTransactionFixJobAgency
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CVPosTransactionFixJobAgency");
            }
        }

        public static string EPaymentsJobHost_CVPosTransactionFixJobEvent
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CVPosTransactionFixJobEvent");
            }
        }

        public static string EPaymentsJobHost_CVPosTransactionFixJobStartTime
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CVPosTransactionFixJobStartTime");
            }
        }

        public static string EPaymentsJobHost_CVPosTransactionFixJobStartDate
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CVPosTransactionFixJobStartDate");
            }
        }

        public static string EPaymentsJobHost_CVPosTransactionFixJobEndDate
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:CVPosTransactionFixJobEndDate");
            }
        }

        //EPaymentsJobHost_UnprocessedVposRequestsJob

        public static bool EPaymentsJobHost_UnprocessedVposRequestsJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:UnprocessedVposRequestsJobEnabled");
            }
        }


        public static int EPaymentsJobHost_UnprocessedVposRequestsJobPeriodInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:UnprocessedVposRequestsJobPeriodInMinutes");
            }
        }

        public static int EPaymentsJobHost_UnprocessedVposRequestsJobFiBankMaxFailedAttempts
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:UnprocessedVposRequestsJobFiBankMaxFailedAttempts");
            }
        }

        public static int EPaymentsJobHost_UnprocessedVposRequestsJobFiBankFailedAttemptTimeoutInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:UnprocessedVposRequestsJobFiBankFailedAttemptTimeoutInMinutes");
            }
        }

        public static int EPaymentsJobHost_UnprocessedVposRequestsJobFibankRequestTimeoutInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:UnprocessedVposRequestsJobFibankRequestTimeoutInMinutes");
            }
        }

        public static string EPaymentsJobHost_UnprocessedVposRequestsJobFiBankVposMerchantHandlerUrl
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:UnprocessedVposRequestsJobFiBankVposMerchantHandlerUrl");
            }
        }

        //DelivertyNotificationJob
        public static int EPaymentsJobHost_EDeliveryNotificationJobPeriodInSeconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EDeliveryficationJobPeriodInSeconds");
            }
        }

        public static bool EPaymentsJobHost_EDeliveryNotificationJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:EDeliveryNotificationJobEnabled");
            }
        }

        public static int EPaymentsJobHost_EDeliveryNotificationJobParallelTasks
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EDeliveryNotificationJobParallelTasks");
            }
        }

        public static int EPaymentsJobHost_EDeliveryNotificationJobBatchSize
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EDeliveryNotificationJobBatchSize");
            }
        }

        public static int PaymentsJobHost_EDeliveryNotificationJobMaxFailedAttempts
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EDeliveryNotificationJobMaxFailedAttempts");
            }
        }

        public static int EPaymentsJobHost_EDeliveryNotificationTimeoutInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:EDeliveryNotificationTimeoutInMinutes");
            }
        }

        public static bool EPaymentsJobHost_EDeliveryNotificationJobValidateSslCert
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:EDeliveryNotificationJobValidateSslCert");
            }
        }

        public static string EPaymentsJobHost_EDeliveryNotificationJobServiceOID
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:EDeliveryNotificationJobServiceOID");
            }
        }

        //DelivertyNotificationJob notifications
        public static string EPaymentsJobHost_PaymentAisClientRequestCreated
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:PaymentAisClientRequestCreated");
            }
        }

        public static string EPaymentsJobHost_PaymentApplicantRequestCreated
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:PaymentApplicantRequestCreated");
            }
        }

        public static string EPaymentsJobHost_PaymentRequestAisClientShared
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:PaymentRequestAisClientShared");
            }
        }

        public static string EPaymentsJobHost_PaymentRequestApplicantClientShared
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:PaymentRequestApplicantClientShared");
            }
        }

        public static string EPaymentsJobHost_PaymentRequestAisClientObligationStatusChanged
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:PaymentRequestAisClientObligationStatusChanged");
            }
        }

        public static string EPaymentsJobHost_PaymentRequestApplicantObligationStatusChanged
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:PaymentRequestApplicantObligationStatusChanged");
            }
        }

        public static string EPaymentsJobHost_PaymentRequestAisClientObligationStatusCanceled
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:PaymentRequestAisClientObligationStatusCanceled");
            }
        }

        public static string EPaymentsJobHost_PaymentRequestApplicantObligationStatusCanceled
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:PaymentRequestApplicantObligationStatusCanceled");
            }
        }

        //Borica Distribution
        public static bool EPaymentsJobHost_DistributionJobEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:DistributionJobEnabled");
            }
        }

        public static int EPaymentsJobHost_DistributionJobPeriodInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:DistributionJobPeriodInMinutes");
            }
        }

        public static int? EPaymentsJobHost_DistributionTransactionsToTake
        {
            get
            {
                return GetAppConfigValue<int?>("EPayments.Job.Host:DistributionTransactionsToTake");
            }
        }

        public static int EPaymentsJobHost_DistributionJobItemsToTake
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:DistributionJobItemsToTake");
            }
        }

        public static string EPaymentsJobHost_DistributionBICCode
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:DistributionBICCode");
            }
        }

        public static string EPaymentsJobHost_DistributionBulstat
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:DistributionBulstat");
            }
        }

        public static string EPaymentsJobHost_DistributionIban
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:DistributionIban");
            }
        }

        public static string EPaymentsJobHost_DistributionSenderName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:DistributionSenderName");
            }
        }

        public static int? EPaymentsJobHost_DistributionJobParentEserviceClient
        {
            get
            {
                return GetAppConfigValue<int?>("EPayments.Job.Host:DistributionJobParentEserviceClient");
            }
        }

        public static string EPaymentsJobHost_SchemasDirectory
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:SchemasDirectory");
            }
        }

        public static string EPaymentsJobHost_XsdFileName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:XsdFileName");
            }
        }

        public static string EPaymentsJobHost_DistributionXmlDirectory
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:DistributionXmlDirectory");
            }
        }

        public static string EPaymentsJobHost_Vpn
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:Vpn");
            }
        }

        public static string EPaymentsJobHost_Vd
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:Vd");
            }
        }

        public static string EPaymentsJobHost_FirstDescription
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:FirstDescription");
            }
        }

        public static string EPaymentsJobHost_SecondDescription
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:SecondDescription");
            }
        }

        public static string EPaymentsJobHost_StartTime
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Job.Host:StartTime");
            }
        }

        // BoricaUnprocessedRequestsJob
        public static int EPaymentsJobHost_BoricaUnprocessedRequestsJobPeriodInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:BoricaUnprocessedRequestsJobPeriodInMinutes");
            }
        }

        public static int EPaymentsJobHost_BoricaUnprocessedRequestsTransactionsToTake
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:BoricaUnprocessedRequestsTransactionsToTake");
            }
        }

        public static int EPaymentsJobHost_BoricaUnprocessedRequestsWaitAfterEachRequestSeconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:BoricaUnprocessedRequestsWaitAfterEachRequestSeconds");
            }
        }

        public static int EPaymentsJobHost_BoricaUnprocessedRequestsLimitTimeSpanInMinutes
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:BoricaUnprocessedRequestsLimitTimeSpanInMinutes");
            }
        }

        public static bool EPaymentsJobHost_BoricaUnprocessedRequestsEnabled
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Job.Host:BoricaUnprocessedRequestsEnabled");
            }
        }

        public static int EPaymentsJobHost_BoricaFinalLimitInHours
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:BoricaFinalLimitInHours");
            }
        }

        public static int EPaymentsJobHost_BoricaRetryPeriodInSeconds
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:BoricaRetryPeriodInSeconds");
            }
        }

        public static int EPaymentsJobHost_BoricaRetryCount
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Job.Host:BoricaRetryCount");
            }
        }

        #endregion

        #region EPayments.Common

        public static string EPaymentsCommon_WebAddress
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Common:WebAddress");
            }
        }

        public static string EPaymentsCommon_BoricaCertificateFolder
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Common:BoricaCertificateFolder");
            }
        }

        public static string EPaymentsCommon_FiBankCertificateFolder
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Common:FiBankCertificateFolder");
            }
        }

        public static bool EPaymentsCommon_UseMachineKeySet
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Common:UseMachineKeySet");
            }
        }

        public static string EPaymentsCommon_EuroExchangeRate
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Common:EuroExchangeRate");
            }
        }

        public static DateTime EPaymentsCommon_EuroAcceptanceDate
        {
            get
            {
                return GetAppConfigValue<DateTime>("EPayments.Common:EuroAcceptanceDate");
            }
        }

        public static string EpaymentsCommon_CurrencyMode
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Common:CurrencyMode");
            }
        }


        public static int EPaymentsCommon_MaxPdfExportEntriesCount
        {
            get
            {
                return GetAppConfigValue<int>("EPayments.Common:MaxPdfExportEntriesCount");
            }
        }

        #endregion

        #region EPayments.EDelivery

        public static string EPaymentsEDelivery_CertificateFolder
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certificates", "EDelivery");
            }
        }

        public static string EPaymentsEDelivery_EDeliveryCertificateName
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:EDeliveryCertificateName");
            }
        }

        public static string EPaymentsEDelivery_PrivateKeyPassphrase
        {
            get
            {
                return GetAppConfigValue<string>("EPayments.Web:PrivateDeliveryKeyPassphrase");
            }
        }

        public static bool EPaymentsEDelivery_PrivateDeliveryUseProduction
        {
            get
            {
                return GetAppConfigValue<bool>("EPayments.Web:PrivateDeliveryUseProduction");
            }
        }
        #endregion

        #region Private

        private static T GetAppConfigValue<T>(string appConfigKey)
        {
            string appConfigValue = System.Configuration.ConfigurationManager.AppSettings[appConfigKey];

            T configValue = (T)TypeDescriptor.GetConverter(typeof(T)).ConvertFromString(appConfigValue);

            return configValue;
        }

        #endregion
    }
}
