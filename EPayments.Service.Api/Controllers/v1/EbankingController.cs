using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using System;
using System.Web.Http;
using EPayments.Common.DataObjects;
using System.Text;
using Newtonsoft.Json.Linq;
using EPayments.Service.Api.Common.CustomExceptions;
using EPayments.Log.ActionLogger;
using EPayments.Service.Api.Auth.Attributes;
using EPayments.Data.ViewObjects.Api;
using System.Collections.Generic;
using System.Linq;
using EPayments.Model.Models;
using System.Web;
using EPayments.Common.Data;

namespace EPayments.Service.Api.Controllers.v1
{
    [EbankingAuthorize]
    [RoutePrefix("api/v1/eBanking")]
    public class EbankingController : ApiController
    {
        private IUnitOfWork unitOfWork;
        private IApiRepository apiRepository;
        private ISystemRepository systemRepository;

        public EbankingController(IUnitOfWork unitOfWork, IApiRepository apiRepository, ISystemRepository systemRepository)
        {
            this.unitOfWork = unitOfWork;
            this.apiRepository = apiRepository;
            this.systemRepository = systemRepository;
        }

        [Route("pendingPaymentsXml")]
        [HttpPost]
        public object PostPendingPaymentsXml(AuthRequestDO requestDO)
        {
            string egn, lnch, bulstat;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                egn = dataJObject.GetValue("egn", StringComparison.OrdinalIgnoreCase).Value<string>();
                lnch = dataJObject.GetValue("lnch", StringComparison.OrdinalIgnoreCase).Value<string>();
                bulstat = dataJObject.GetValue("bulstat", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            UinType applicantUinTypeId;
            string applicantUin;

            if (!String.IsNullOrWhiteSpace(egn) && String.IsNullOrWhiteSpace(lnch) && String.IsNullOrWhiteSpace(bulstat))
            {
                applicantUinTypeId = UinType.Egn;
                applicantUin = egn;
            }
            else if (String.IsNullOrWhiteSpace(egn) && !String.IsNullOrWhiteSpace(lnch) && String.IsNullOrWhiteSpace(bulstat))
            {
                applicantUinTypeId = UinType.Lnch;
                applicantUin = lnch;
            }
            else if (String.IsNullOrWhiteSpace(egn) && String.IsNullOrWhiteSpace(lnch) && !String.IsNullOrWhiteSpace(bulstat))
            {
                applicantUinTypeId = UinType.Bulstat;
                applicantUin = bulstat;
            }
            else
            {
                throw new CustomServiceException("Invalid post data. Only one identifier between 'egn', 'lnch' and 'bulstat' must be specified.");
            }

            var pendingRequestTuples = this.apiRepository.GetRequestInfosByApplicant(applicantUinTypeId, applicantUin);

            if (pendingRequestTuples.Count > 0)
            {
                LogEbankingAccess(requestDO.ClientId, pendingRequestTuples.Select(e => e.Item1).ToList());
            }

            return new
            {
                PaymentRequests = pendingRequestTuples.Select(e => e.Item2).ToList()
            };
        }

        [Route("pendingPaymentsJson")]
        [HttpPost]
        public object PostPendingPaymentsJson(AuthRequestDO requestDO)
        {
            string egn, lnch, bulstat;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                egn = dataJObject.GetValue("egn", StringComparison.OrdinalIgnoreCase).Value<string>();
                lnch = dataJObject.GetValue("lnch", StringComparison.OrdinalIgnoreCase).Value<string>();
                bulstat = dataJObject.GetValue("bulstat", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            UinType applicantUinTypeId;
            string applicantUin;

            if (!String.IsNullOrWhiteSpace(egn) && String.IsNullOrWhiteSpace(lnch) && String.IsNullOrWhiteSpace(bulstat))
            {
                applicantUinTypeId = UinType.Egn;
                applicantUin = egn;
            }
            else if (String.IsNullOrWhiteSpace(egn) && !String.IsNullOrWhiteSpace(lnch) && String.IsNullOrWhiteSpace(bulstat))
            {
                applicantUinTypeId = UinType.Lnch;
                applicantUin = lnch;
            }
            else if (String.IsNullOrWhiteSpace(egn) && String.IsNullOrWhiteSpace(lnch) && !String.IsNullOrWhiteSpace(bulstat))
            {
                applicantUinTypeId = UinType.Bulstat;
                applicantUin = bulstat;
            }
            else
            {
                throw new CustomServiceException("Invalid post data. Only one identifier between 'egn', 'lnch' and 'bulstat' must be specified.");
            }

            var pendingRequestTuples = this.apiRepository.GetParsedRequestInfosByApplicant(applicantUinTypeId, applicantUin);

            if (pendingRequestTuples.Count > 0)
            {
                LogEbankingAccess(requestDO.ClientId, pendingRequestTuples.Select(e => e.Item1).ToList());
            }

            return new
            {
                PaymentRequests = pendingRequestTuples.Select(e => e.Item2).ToList()
            };
        }

        [Route("paymentByAccessCodeXml")]
        [HttpPost]
        public object PostPaymentByAccessCodeXml(AuthRequestDO requestDO)
        {
            string accessCode;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                accessCode = dataJObject.GetValue("accessCode", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var requestTuple = this.apiRepository.GetRequestInfoByAccessCode(accessCode);

            if (requestTuple != null)
            {
                LogEbankingAccess(requestDO.ClientId, requestTuple.Item1);
            }

            return new
            {
                Id = requestTuple != null ? requestTuple.Item2.Id : null,
                Status = requestTuple != null ? requestTuple.Item2.Status : (PaymentRequestStatus?)null,
                ChangeTime = requestTuple != null ? requestTuple.Item2.ChangeTime : (DateTime?)null,
                RequestXml = requestTuple != null ? requestTuple.Item2.RequestXml : null,
            };
        }

        [Route("paymentByAccessCodeJson")]
        [HttpPost]
        public object PostPaymentByAccessCodeJson(AuthRequestDO requestDO)
        {
            string accessCode;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                accessCode = dataJObject.GetValue("accessCode", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var requestTuple = this.apiRepository.GetParsedRequestInfoByAccessCode(accessCode);

            if (requestTuple != null)
            {
                LogEbankingAccess(requestDO.ClientId, requestTuple.Item1);
            }

            RequestInfoParsedVO returnValue = null;

            if (requestTuple == null)
            {
                returnValue = new RequestInfoParsedVO();
            }
            else
            {
                returnValue = requestTuple.Item2;
            }

            return returnValue;
        }

        private void LogEbankingAccess(string clientId, int paymentRequestId)
        {
            LogEbankingAccess(clientId, new List<int>() { paymentRequestId });
        }

        private void LogEbankingAccess(string clientId, List<int> paymentRequestIds)
        {
            if (paymentRequestIds != null && paymentRequestIds.Count > 0)
            {
                EbankingClient еbankingClient = this.systemRepository.GetEbankingClientByClientId(clientId);

                foreach(int paymentRequestId in paymentRequestIds)
                {
                    EbankingAccessLog accessLog = new EbankingAccessLog();
                    accessLog.EbankingClientId = еbankingClient.EbankingClientId;
                    accessLog.PaymentRequestId = paymentRequestId;
                    accessLog.IpAddress = HttpContext.Current.Request.UserHostAddress;
                    accessLog.AccessDate = DateTime.Now;

                    this.systemRepository.AddEntity<EbankingAccessLog>(accessLog);
                }

                this.unitOfWork.Save();
            }
        }
    }
}
