using EPayments.Common;
using EPayments.CVPosTransaction.CVPosService;
using EPayments.Model.Enums;
using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.Threading.Tasks;

namespace EPayments.CVPosTransaction.Manager
{
    public class CVPosRegisterManager: ICVPosRegisterManager
    {
        
        private NapWSClient RegisterService { get; set; }
        private string Message { get; set; }
 
        public CVPosRegisterManager()
           : this("NapWSPort")
        { }

        public CVPosRegisterManager(string bidingName)
        {
            this.RegisterService = new NapWSClient(bidingName);

            try
            {
                X509KeyStorageFlags flags = AppSettings.EPaymentsCommon_UseMachineKeySet == false ?
                        X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable :
                        X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable;

                var cert = new X509Certificate2(
                    Path.Combine(AppSettings.EPaymentsWeb_BoricaCentralVposCertificateFolder,
                    AppSettings.EPaymentsJobHost_CentralVposPrivateKeyFileName),
                    AppSettings.EPaymentJobHost_CentralVposPrivateKeyPassphrase,
                    flags);

                this.RegisterService.ClientCredentials.ClientCertificate.Certificate = cert;
            }
            catch (Exception ex)
            {
                var newEx = new Exception(this.Message, ex);
                newEx.Data.Add("Type", CVPosExeptionEnums.Communication);
                newEx.Data.Add("Message", "Проблем със сертификата");
                newEx.Data.Add("Code", "");
                throw newEx;
            }
        }

        public void DistributionRevenueAgencies(string agency, DateTime distributedDate, int num,
            recTransaction[] transactions, decimal totalSum, recDistributedAmount[] distributedAmount)
        {
            try
            {
                AddTls();

                this.RegisterService.distributionRevenueAgencies(agency, distributedDate,num,transactions,
                    totalSum, distributedAmount);
            }
            catch (FaultException)
            {
                throw;
            }
            catch (EndpointNotFoundException)
            {
                throw;
            }
            catch (CommunicationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                this.Message = "EventRegisterManager LogEventAsync exception: "
                             + (ex.Message ?? String.Empty) + (ex.InnerException != null
                             ? " Inner exception: " +
                             ex.InnerException.Message
                             ?? String.Empty : String.Empty);
                
                throw new Exception(this.Message, ex);
            }
        }

        public recEventTransaction[] GetTransactionPerDate(string agency, string @event, DateTime dateEvent, string tid)
        {
            try
            {
                AddTls();

                return this.RegisterService.transactionsForDate(agency, @event, dateEvent, tid);
            }
            catch (Exception ex)
            {

                this.Message = "EventRegisterManager LogEventAsync exception: "
                            + (ex.Message ?? String.Empty) + (ex.InnerException != null
                            ? " Inner exception: " +
                            ex.InnerException.Message
                            ?? String.Empty : String.Empty);

                switch (ex.GetType().BaseType.Name)
                {
                    case nameof(FaultException):
                        {
                            FaultException faultException = (FaultException)ex;
                            var newEx = new Exception(this.Message, ex);
                            newEx.Data.Add("Type", CVPosExeptionEnums.Fault);
                            newEx.Data.Add("Message", faultException.Message);
                            newEx.Data.Add("Code", faultException.Code.Name);
                            throw newEx;
                        }
                }

                switch (ex.GetType().Name)
                {
                    case nameof(EndpointNotFoundException):
                        {
                            var newEx = new Exception(this.Message, ex);
                            newEx.Data.Add("Type", CVPosExeptionEnums.EndpointNotFound);
                            newEx.Data.Add("Message", ex.Message);
                            newEx.Data.Add("Code", string.Empty);
                            throw newEx;
                        }                        
                   case nameof(CommunicationException):
                        {
                            var newEx = new Exception(this.Message, ex);
                            newEx.Data.Add("Type", CVPosExeptionEnums.Communication);
                            newEx.Data.Add("Message", ex.Message);
                            newEx.Data.Add("Code", string.Empty);
                            throw newEx;
                        }
                   
                    default:
                        throw new Exception(this.Message, ex);
                }
            }
        }


        public async Task<transactionsForDateResponse> GetTransactionPerDateAsync (string agency, string @event, DateTime dateEvent, string tid)
        {
            try
            {
                AddTls();

                return await this.RegisterService.transactionsForDateAsync(agency, @event, dateEvent, tid);
            }
            catch (Exception ex)
            {

                this.Message = "EventRegisterManager LogEventAsync exception: "
                            + (ex.Message ?? String.Empty) + (ex.InnerException != null
                            ? " Inner exception: " +
                            ex.InnerException.Message
                            ?? String.Empty : String.Empty);

                switch (ex.GetType().BaseType.Name)
                {
                    case nameof(FaultException):
                        {
                            FaultException faultException = (FaultException)ex;
                            var newEx = new Exception(this.Message, ex);
                            newEx.Data.Add("Type", CVPosExeptionEnums.Fault);
                            newEx.Data.Add("Message", faultException.Message);
                            newEx.Data.Add("Code", faultException.Code.Name);
                            throw newEx;
                        }
                }

                switch (ex.GetType().Name)
                {
                    case nameof(EndpointNotFoundException):
                        {
                            var newEx = new Exception(this.Message, ex);
                            newEx.Data.Add("Type", CVPosExeptionEnums.EndpointNotFound);
                            newEx.Data.Add("Message", ex.Message);
                            newEx.Data.Add("Code", string.Empty);
                            throw newEx;
                        }
                    case nameof(CommunicationException):
                        {
                            var newEx = new Exception(this.Message, ex);
                            newEx.Data.Add("Type", CVPosExeptionEnums.Communication);
                            newEx.Data.Add("Message", ex.Message);
                            newEx.Data.Add("Code", string.Empty);
                            throw newEx;
                        }

                    default:
                        throw new Exception(this.Message, ex);
                }
            }
        }

        private void AddTls()
        {
            if ((System.Net.ServicePointManager.SecurityProtocol & System.Net.SecurityProtocolType.Tls12) != System.Net.SecurityProtocolType.Tls12)
            {
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            }
        }
    }
}
