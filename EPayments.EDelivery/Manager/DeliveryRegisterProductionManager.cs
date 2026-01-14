using EPayments.Common;
using EPayments.EDelivery.EDeliveryProductionClient;
using EPayments.EDelivery.Models;
using EPayments.Model.Models;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace EPayments.EDelivery.Manager
{
    public class DeliveryRegisterProductionManager : DeliveryRegisterBaseManager, IDeliveryRegisterManager
    {
        public DeliveryRegisterProductionManager()
           : this("BasicHttpBinding_IEDeliveryIntegrationService2")
        { }
        
        public DeliveryRegisterProductionManager(string bidingName)
        {
            this.RegisterService = new EDeliveryIntegrationServiceClient(bidingName);

            LoadConfiguration();

            if (!ValidateSslCert == true)
            {
                try
                {
                    X509KeyStorageFlags flags = AppSettings.EPaymentsCommon_UseMachineKeySet == false ?
                        X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable :
                        X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable;

                    var cert = new X509Certificate2(Path.Combine(AppSettings.EPaymentsEDelivery_CertificateFolder, AppSettings.EPaymentsEDelivery_EDeliveryCertificateName),
                     AppSettings.EPaymentsEDelivery_PrivateKeyPassphrase,
                     flags);

                    this.RegisterService.ClientCredentials.ClientCertificate.Certificate = cert;
                }
                catch (Exception ex)
                {
                    this.Message = "EventRegisterManager LogEventAsync exception: " + (ex.Message ?? String.Empty) + (ex.InnerException != null ? " Inner exception: " + ex.InnerException.Message ?? String.Empty : String.Empty);
                    throw new Exception(this.Message, ex);
                }
            }
        }

        private EDeliveryIntegrationServiceClient RegisterService { get; set; }
        private string ServiceOID { get; set; }
        private string SystemOID { get; set; }
        private string Message { get; set; }

        private bool ValidateSslCert;
        private bool IsReady { get; set; }

        public DepartmentInstitutionInfo[] GetAdministration()
        {
            try
            {
                if (base.ContainsDepartments())
                {
                    DepartmentInstitutionInfo[] mappedDepartments = base.GetDepartmentsFromCache();

                    if (mappedDepartments != null)
                    {
                        return mappedDepartments;
                    }
                }
                
                DcInstitutionInfo[] departments = this.RegisterService.GetRegisteredInstitutionsAsync().Result;

                return Map(departments);
            }
            catch (Exception ex)
            {
                this.Message = "EventRegisterManager LogEventAsync exception: "
                    + (ex.Message ?? String.Empty) +
                    (ex.InnerException != null ? " Inner exception: " + ex.InnerException.Message
                    ?? String.Empty : String.Empty);

                throw new Exception(this.Message, ex);
            }
        }

        public async Task<DepartmentInstitutionInfo[]> GetAdministrationAsync()
        {
            try
            {
                if (base.ContainsDepartments())
                {
                    DepartmentInstitutionInfo[] mappedDepartments = base.GetDepartmentsFromCache();

                    if (mappedDepartments != null)
                    {
                        return mappedDepartments;
                    }
                }

                DcInstitutionInfo[] departments = await this.RegisterService.GetRegisteredInstitutionsAsync();

                return Map(departments);
            }
            catch (Exception ex)
            {
                this.Message = "EventRegisterManager LogEventAsync exception: "
                    + (ex.Message ?? String.Empty) +
                    (ex.InnerException != null ? " Inner exception: " + ex.InnerException.Message
                    ?? String.Empty : String.Empty);
                throw new Exception(this.Message, ex);
            }
        }

        public async Task<int> SendMessageAsync(EserviceDeliveryNotification notification, string message)
        {
            return await this.RegisterService.SendMessageAsync(new DcMessageDetails()
            {
                Title = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobSPName,
                MessageText = message,
                AttachedDocuments = new DcDocument[] { }
            },
            eProfileType.Institution,
            notification.Uniqueidentifier.Trim(),
            null,
            null,
            serviceOID: this.ServiceOID,
            operatorEGN: null);
        }

        public async Task<int> SendMessagePersonAsync(EserviceDeliveryNotification notification, string message, bool isInstitution = false)
        {
            return await this.RegisterService.SendMessageAsync(new DcMessageDetails()
            {
                Title = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobSPName,
                MessageText = message,
                AttachedDocuments = new DcDocument[] { }
            },
            isInstitution == false ? eProfileType.Person : eProfileType.LegalPerson,
            notification.PersonUniqueIdentifier.Trim(),
            null,
            null,
            serviceOID: this.ServiceOID,
            operatorEGN: null);
        }

        private void LoadConfiguration()
        {
            try
            {
                this.ServiceOID = AppSettings.EPaymentsJobHost_EDeliveryNotificationJobServiceOID;
                this.SystemOID = AppSettings.EPaymentsJobHost_EDeliveryNotificationJobInformationSystemOID;
                this.ValidateSslCert = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobValidateSslCert;
                this.IsReady = true;
            }
            catch
            {
                this.IsReady = false;
            }
        }

        private DepartmentInstitutionInfo[] Map(DcInstitutionInfo[] departments)
        {
            DepartmentInstitutionInfo[] mappedDepartments = departments.Select(d => new DepartmentInstitutionInfo()
            {
                HeadInstitution = DepartmentSubjectPublicInfo.Map(d.HeadInstitution),
                Name = d.Name,
                SubInstitutions = d.SubInstitutions.Select(di => DepartmentSubjectPublicInfo.Map(di)).ToArray(),
                ExtensionData = d.ExtensionData,
                ElectronicSubjectId = d.ElectronicSubjectId,
                ElectronicSubjectName = d.ElectronicSubjectName,
                Email = d.Email,
                IsActivated = d.IsActivated,
                PhoneNumber = d.PhoneNumber,
                ProfileType = (DepartmentProfileTypeEnum)((int)d.ProfileType),
                DateCreated = null,
                UniqueSubjectIdentifier = d.UniqueSubjectIdentifier,
                VerificationInfo = d.VerificationInfo
            }).ToArray();

            base.AddDepartmentsToCache(mappedDepartments);

            return mappedDepartments;
        }
    }
}
