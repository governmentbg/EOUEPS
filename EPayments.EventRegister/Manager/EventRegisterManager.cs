using EPayments.EventRegister.DataObjects;
using EPayments.EventRegister.EventRegisterService;
using EPayments.Common;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.EventRegister.Manager
{
    public class EventRegisterManager : IEventRegisterManager
    {
        //private static ILog logger = LogManager.GetLogger("JournalLogger");
        private RequestedServiceType Service { get; set; }
        private RegisterEventSOAPPortClient EventRegisterService  { get; set; }
        private string ServiceOIDDefault { get; set; }
        private string SystemOID { get; set; }
        private bool ValidateSslCert;

        public EventRegisterManager()
            : this("RegisterEventSOAPPortImplPort")
        { }

        public EventRegisterManager(string bidingName)
        {
            this.EventRegisterService = new RegisterEventSOAPPortClient(bidingName);
            LoadConfiguration();
            if (!ValidateSslCert)
            {
                ServicePointManager.ServerCertificateValidationCallback += delegate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };
            }
        }

        public void LogEvent(EventDO ev)
        {
            EAServiceLCEventType wsEv = CreateEAServiceLCEventType(ev);

            this.EventRegisterService.RegisterEvent(new EAServiceLCEventType[] { wsEv });
        }

        public async Task LogEventAsync(EventDO ev)
        {
            EAServiceLCEventType wsEv = CreateEAServiceLCEventType(ev);

            await this.EventRegisterService.RegisterEventAsync(new EAServiceLCEventType[] { wsEv });
        }

        public bool IsReady { get; private set; }

        private EAServiceLCEventType CreateEAServiceLCEventType(EventDO ev)
        {
            EAServiceLCEventType wsEv = new EAServiceLCEventType();

            wsEv.EventTime = ev.Time;
            wsEv.EventType1 = ev.EventType;
            wsEv.EventDescription = ev.EventDescription;
            wsEv.DocumentRegId = ev.DocumentRegNumber;
            wsEv.InformationSystemOID = this.SystemOID;
            wsEv.RequestedService = this.Service;
            wsEv.RequestedService.ServiceOID = !String.IsNullOrEmpty(ev.ServiceOID) ? ev.ServiceOID : ServiceOIDDefault;

            return wsEv;
        }

        private void LoadConfiguration()
        {
            try
            {
                this.Service = new RequestedServiceType();
                this.Service.AdminiOID = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobAdminOID;
                this.Service.AdminLegalName = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobAdminLegalName;
                this.Service.ServiceOID = ServiceOIDDefault = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobServiceOID;
                this.Service.SPOID = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobSPOID;
                this.Service.SPName = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobSPName;
                this.SystemOID = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobInformationSystemOID;

                this.ValidateSslCert = AppSettings.EPaymentsJobHost_EventRegisterNotificationJobValidateSslCert;

                //logger.Info("Journal service configured successfully!");
                this.IsReady = true;
            }
            catch
            {
                //logger.Info("Journal logger is disabled!");
                //logger.Error("Can not load configuration for journal logger!", ex);
                this.IsReady = false;
            }
        }
    }
}
