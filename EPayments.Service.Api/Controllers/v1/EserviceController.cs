using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Documents.Serializer;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using EPayments.Service.Api.Common;
using System;
using System.Collections.Generic;
using System.Web.Http;
using System.Linq;
using EPayments.Common.Helpers;
using System.Text;
using Newtonsoft.Json.Linq;
using EPayments.Service.Api.Common.CustomExceptions;
using EPayments.Service.Api.Common.XsdValidator;
using EPayments.Service.Api.Auth.Attributes;
using EPayments.Common.DataObjects;
using EPayments.Common;
using EPayments.Model.DataObjects.EmailTemplateContext;
using Newtonsoft.Json;
using EPayments.Service.Api.Controllers.v1.DataObjects;
using EPayments.Data.ViewObjects.Api;
using System.Text.RegularExpressions;

namespace EPayments.Service.Api.Controllers.v1
{
    [EserviceAuthorize]
    [RoutePrefix("api/v1/eService")]
    public class EserviceController : ApiController
    {
        private IUnitOfWork unitOfWork;
        private IDocumentSerializer documentSerializer;
        private IApiRepository apiRepository;
        private ICommonRepository commonRepository;
        private ISystemRepository systemRepository;
        private ISchemaValidator schemaValidator;
        private IPaymentRequestRepository paymentRequestRepository;

        public EserviceController(
            IUnitOfWork unitOfWork,
            IDocumentSerializer documentSerializer,
            IApiRepository apiRepository,
            ICommonRepository commonRepository,
            ISystemRepository systemRepository,
            ISchemaValidator schemaValidator,
            IPaymentRequestRepository paymentRequestRepository)
        {
            this.unitOfWork = unitOfWork;
            this.documentSerializer = documentSerializer;
            this.apiRepository = apiRepository;
            this.commonRepository = commonRepository;
            this.systemRepository = systemRepository;
            this.schemaValidator = schemaValidator;
            this.paymentRequestRepository = paymentRequestRepository;
        }

        [Route("payment")]
        [HttpPost]
        public object PostPayment(AuthRequestDO requestDO)
        {
            string requestXmlContent = null;
            string aisPaymentId = null;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                requestXmlContent = dataJObject.GetValue("requestXml", StringComparison.OrdinalIgnoreCase).Value<string>();
                var aisPaymentIdObj = dataJObject.GetValue("aisPaymentId", StringComparison.OrdinalIgnoreCase);
                if (aisPaymentIdObj != null)
                    aisPaymentId = aisPaymentIdObj.Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            Tuple<R_10052.PaymentRequestUnacceptedReceipt, R_10055.PaymentRequestAcceptedReceipt> receiptsTuple =
                InternalPostPayment(requestDO, requestXmlContent, aisPaymentId);

            string unacceptedReceiptXml = receiptsTuple.Item1 != null ? this.documentSerializer.XmlSerializeToString(receiptsTuple.Item1) : null;
            string acceptedReceiptXml = receiptsTuple.Item2 != null ? this.documentSerializer.XmlSerializeToString(receiptsTuple.Item2) : null;

            return new
            {
                UnacceptedReceiptXml = unacceptedReceiptXml,
                AcceptedReceiptXml = acceptedReceiptXml
            };
        }

        [Route("paymentJson")]
        [HttpPost]
        public object PostPaymentJson(AuthRequestDO requestDO)
        {
            PaymentRequestJsonDO paymentRequestJson;
            string requestXmlContent = null;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                paymentRequestJson = JsonConvert.DeserializeObject<PaymentRequestJsonDO>(dataJson);

                EserviceClient еserviceClient = this.systemRepository.GetEserviceClientByClientId(requestDO.ClientId);
                if (еserviceClient.AccountIBAN != paymentRequestJson.ServiceProviderIBAN)
                {
                    return BadRequest("Requested IBAN must match the AIS client IBAN");
                }

                R_10046.PaymentRequest paymentRequestRio = CreatePaymentRequestRioFromJson(paymentRequestJson);
                requestXmlContent = documentSerializer.XmlSerializeToString(paymentRequestRio);
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            Tuple<R_10052.PaymentRequestUnacceptedReceipt, R_10055.PaymentRequestAcceptedReceipt> receiptsTuple =
                InternalPostPayment(requestDO, requestXmlContent, paymentRequestJson.AisPaymentId);

            UnacceptedReceiptJsonDO unacceptedReceiptJsonDO = receiptsTuple.Item1 != null ? CreateUnacceptedReceiptJsonDO(receiptsTuple.Item1) : null;
            AcceptedReceiptJsonDO acceptedReceiptJsonDO = receiptsTuple.Item2 != null ? CreateAcceptedReceiptJsonDO(receiptsTuple.Item2) : null;

            return new
            {
                UnacceptedReceiptJson = unacceptedReceiptJsonDO,
                AcceptedReceiptJson = acceptedReceiptJsonDO
            };
        }

        [Route("paymentJsonExtended")]
        [HttpPost]
        public object PostPaymentJsonExtended(AuthRequestExtendedDO requestDO)
        {
            if (string.IsNullOrEmpty(requestDO.EserviceClientId))
            {
                return BadRequest("EserviceClientId is mandatory");
            }

            PaymentRequestJsonExtendedDO paymentRequestJson;
            string requestXmlContent = null;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                paymentRequestJson = JsonConvert.DeserializeObject<PaymentRequestJsonExtendedDO>(dataJson);
                paymentRequestJson.AdditionalInformation = JsonHelper.RepairBrokenAdditionalInformationJsonString(paymentRequestJson.AdditionalInformation);
            }
            catch
            {
                return BadRequest("request body can not be deserialize, please take a look of your request data parameters");
            }

            R_10046.PaymentRequest paymentRequestRio = CreatePaymentRequestRioFromJsonExtendet(paymentRequestJson, requestDO.EserviceClientId);
            requestXmlContent = documentSerializer.XmlSerializeToString(paymentRequestRio);

            Tuple<R_10052.PaymentRequestUnacceptedReceipt, R_10055.PaymentRequestAcceptedReceipt> receiptsTuple =
                InternalPostPayment(requestDO, requestXmlContent, paymentRequestJson.AisPaymentId, true, requestDO.EserviceClientId, paymentRequestJson.PayOrder);

            UnacceptedReceiptJsonDO unacceptedReceiptJsonDO = receiptsTuple.Item1 != null ? CreateUnacceptedReceiptJsonDO(receiptsTuple.Item1) : null;
            AcceptedReceiptJsonDO acceptedReceiptJsonDO = receiptsTuple.Item2 != null ? CreateAcceptedReceiptJsonDO(receiptsTuple.Item2) : null;

            var accessCode = string.Empty;
            if (acceptedReceiptJsonDO != null) 
            {
                var paymentRequest = apiRepository.GetRequestXmlsByIndetifiers(new List<string>() { acceptedReceiptJsonDO.Id }).FirstOrDefault().PaymentRequest;
                accessCode = paymentRequest.PaymentRequestAccessCode;
            }

            AcceptedReceiptExtendedJsonDO acceptedReceiptExtendedJsonDO = new AcceptedReceiptExtendedJsonDO()
            {
                Id = acceptedReceiptJsonDO?.Id,
                RegistrationTime = acceptedReceiptJsonDO == null ? DateTime.Now : acceptedReceiptJsonDO.RegistrationTime,
                AccessCode = accessCode,
            };

            return new
            {
                UnacceptedReceiptJson = unacceptedReceiptJsonDO,
                AcceptedReceiptJson = acceptedReceiptExtendedJsonDO
            };
        }

        [Route("getClientsByEik")]
        [HttpPost]
        public List<RequestEikInfoVO> GetClientsByEik(AuthRequestDO requestDO)
        {
            string eik;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                eik = dataJObject.GetValue("eik", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch (Exception ex)
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            return this.apiRepository.GetServiceClientByEik(eik);
        }

        [Route("multiplePayment")]
        [HttpPost]
        public object PostMultiplePayment(AuthRequestDO requestDO)
        {
            throw new NotImplementedException();
        }

        [Route("paymentsById")]
        [HttpPost]
        public object PostPaymentsById(AuthRequestDO requestDO)
        {
            List<string> requestIds;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                requestIds = dataJObject.GetValue("requestIds", StringComparison.OrdinalIgnoreCase).Select(e => e.Value<string>()).ToList();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            GetEserviceAuthorizer(requestDO.ClientId).PermitAccessToRequests(requestIds);

            var foundRequests = this.apiRepository.GetRequestXmlsByIndetifiers(requestIds);

            var paymentRequests = new List<object>();
            foreach (var id in requestIds)
            {
                string requestXml = null;
                if (foundRequests.Any(e => e.Id == id))
                {
                    requestXml = foundRequests.Single(e => e.Id == id).RequestXml;
                }

                paymentRequests.Add(new
                {
                    Id = id,
                    RequestXml = requestXml
                });
            }

            return new
            {
                PaymentRequests = paymentRequests
            };
        }

        [Route("paymentsByIdJson")]
        [HttpPost]
        public object PostPaymentsByIdJson(AuthRequestDO requestDO)
        {
            List<string> requestIds;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                requestIds = dataJObject.GetValue("requestIds", StringComparison.OrdinalIgnoreCase).Select(e => e.Value<string>()).ToList();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            GetEserviceAuthorizer(requestDO.ClientId).PermitAccessToRequests(requestIds, authorizeIfClientIdIsPaymentInitiator: true);

            var foundRequests = this.apiRepository.GetRequestXmlsByIndetifiers(requestIds);

            var paymentRequests = new List<object>();
            foreach (var id in requestIds)
            {
                PaymentRequest request = null;
                if (foundRequests.Any(e => e.Id == id))
                {
                    request = foundRequests.Single(e => e.Id == id).PaymentRequest;
                }

                paymentRequests.Add(new
                {
                    Id = id,
                    requestJson = CreatePaymentRequestJsonDO(request)
                });
            }

            return new
            {
                PaymentRequests = paymentRequests
            };
        }

        [Route("paymentsStatus")]
        [HttpPost]
        public object PostPaymentsStatus(AuthRequestDO requestDO)
        {
            List<string> requestIds;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                requestIds = dataJObject.GetValue("requestIds", StringComparison.OrdinalIgnoreCase).Select(e => e.Value<string>()).ToList();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            GetEserviceAuthorizer(requestDO.ClientId).PermitAccessToRequests(requestIds, authorizeIfClientIdIsPaymentInitiator: true);

            var foundRequests = this.apiRepository.GetRequestStatusesByIdentifiers(requestIds);

            var paymentStatuses = new List<object>();
            foreach (var id in requestIds)
            {
                PaymentRequestStatus? status = null;
                DateTime? changeTime = null;

                if (foundRequests.Any(e => e.Id == id))
                {
                    var request = foundRequests.Single(e => e.Id == id);
                    status = request.Status;
                    changeTime = request.ChangeTime;
                }

                paymentStatuses.Add(new
                {
                    Id = id,
                    Status = status,
                    ChangeTime = changeTime
                });
            }

            return new
            {
                PaymentStatuses = paymentStatuses
            };
        }

        [Route("accessCode")]
        [HttpPost]
        public object PostAccessCode(AuthRequestDO requestDO)
        {
            string id;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                id = dataJObject.GetValue("id", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            GetEserviceAuthorizer(requestDO.ClientId).PermitAccessToRequest(id);

            string accessCode = null;

            var paymentRequest = this.commonRepository.FindPaymentRequestByIdentifier(id);

            if (paymentRequest != null)
            {
                if (paymentRequest.PaymentRequestAccessCode == null)
                {
                    paymentRequest.PaymentRequestAccessCode = this.commonRepository.GeneratePaymentRequestAccessCode();

                    this.unitOfWork.Save();
                }

                accessCode = paymentRequest.PaymentRequestAccessCode;
            }

            return new { AccessCode = accessCode };
        }

        [Route("suspendRequest")]
        [HttpPost]
        public object PostSuspendRequest(AuthRequestDO requestDO)
        {
            string id;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                id = dataJObject.GetValue("id", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            GetEserviceAuthorizer(requestDO.ClientId).PermitAccessToRequest(id);

            var paymentRequest = this.commonRepository.FindPaymentRequestByIdentifier(id);

            if (paymentRequest != null)
            {
                if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Pending || paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Canceled)
                {
                    paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Suspended;
                    paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;

                    if (paymentRequest.ObligationStatusId != ObligationStatusEnum.Canceled)
                    {
                        paymentRequest.ObligationStatusId = ObligationStatusEnum.Canceled;
                    }

                    this.unitOfWork.Save();

                    User user = this.systemRepository.GetUserByUin(paymentRequest.ApplicantUin);
                    if (user != null && user.StatusNotifications && !String.IsNullOrWhiteSpace(user.Email))
                    {
                        StatusChangedPaymentRequestContextDO contextDO = new StatusChangedPaymentRequestContextDO(
                            paymentRequest.PaymentRequestIdentifier,
                            paymentRequest.ServiceProviderName,
                            paymentRequest.PaymentReason,
                            paymentRequest.PaymentAmount,
                            paymentRequest.PaymentRequestStatusId.GetDescription(),
                            paymentRequest.CreateDate);

                        Email email = new Email(contextDO, user.Email);

                        this.systemRepository.AddEntity<Email>(email);
                        this.unitOfWork.Save();
                    }

                    if (user != null && user.StatusObligationNotifications && !String.IsNullOrWhiteSpace(user.Email))
                    {
                        StatusChangedObligationContextDO contextOblDO = new StatusChangedObligationContextDO(
                            paymentRequest.PaymentRequestIdentifier,
                            paymentRequest.ServiceProviderName,
                            paymentRequest.PaymentReason,
                            paymentRequest.PaymentAmount,
                            paymentRequest.ObligationStatusId.GetDescription(),
                            paymentRequest.CreateDate);

                        Email email = new Email(contextOblDO, user.Email);

                        this.systemRepository.AddEntity<Email>(email);
                        this.unitOfWork.Save();
                    }

                    if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                    {
                        EserviceNotification statusNotification = new EserviceNotification(paymentRequest);

                        this.systemRepository.AddEntity<EserviceNotification>(statusNotification);
                        this.unitOfWork.Save();
                    }

                    return new object();
                }
                else
                {
                    throw new CustomServiceException("Unable to suspend payment request with status different from 'Pending'.");
                }
            }
            else
            {
                throw new CustomServiceException("Payment request not found.");
            }
        }

        [Route("setStatusPaid")]
        [HttpPost]
        public object PostSetStatusPaid(AuthRequestDO requestDO)
        {
            string id;
            PaidStatusPaymentMethod? paidStatusPaymentMethodId;
            string paidStatusPaymentMethodDescription;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                id = dataJObject.GetValue("id", StringComparison.OrdinalIgnoreCase).Value<string>();
                string paymentMethod = dataJObject.GetValue("paymentMethod", StringComparison.OrdinalIgnoreCase).Value<string>();
                paidStatusPaymentMethodId = (PaidStatusPaymentMethod)Enum.Parse(typeof(PaidStatusPaymentMethod), paymentMethod.Trim(), true);
                paidStatusPaymentMethodDescription = dataJObject.GetValue("paymentDescription", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            GetEserviceAuthorizer(requestDO.ClientId).PermitAccessToRequest(id);

            var paymentRequest = this.commonRepository.FindPaymentRequestByIdentifier(id);

            if (paymentRequest != null)
            {
                if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Pending)
                {
                    paymentRequest.PaymentRequestStatusId = PaymentRequestStatus.Paid;
                    paymentRequest.PaidStatusPaymentMethodId = paidStatusPaymentMethodId.Value;
                    paymentRequest.PaidStatusPaymentMethodDescription = paidStatusPaymentMethodDescription;
                    paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;
                    paymentRequest.ObligationStatusId = Model.Enums.ObligationStatusEnum.IrrevocableOrder;

                    this.unitOfWork.Save();

                    User user = this.systemRepository.GetUserByUin(paymentRequest.ApplicantUin);
                    if (user != null && user.StatusNotifications && !String.IsNullOrWhiteSpace(user.Email))
                    {
                        StatusChangedPaymentRequestContextDO contextDO = new StatusChangedPaymentRequestContextDO(
                            paymentRequest.PaymentRequestIdentifier,
                            paymentRequest.ServiceProviderName,
                            paymentRequest.PaymentReason,
                            paymentRequest.PaymentAmount,
                            paymentRequest.PaymentRequestStatusId.GetDescription(),
                            paymentRequest.CreateDate);

                        Email email = new Email(contextDO, user.Email);

                        this.systemRepository.AddEntity<Email>(email);
                        this.unitOfWork.Save();
                    }
                    if (user != null && user.StatusObligationNotifications && !String.IsNullOrWhiteSpace(user.Email))
                    {
                        StatusChangedObligationContextDO contextOblDO = new StatusChangedObligationContextDO(
                            paymentRequest.PaymentRequestIdentifier,
                            paymentRequest.ServiceProviderName,
                            paymentRequest.PaymentReason,
                            paymentRequest.PaymentAmount,
                            paymentRequest.ObligationStatusId.GetDescription(),
                            paymentRequest.CreateDate);

                        Email email = new Email(contextOblDO, user.Email);

                        this.systemRepository.AddEntity<Email>(email);
                        this.unitOfWork.Save();
                    }

                    if (!String.IsNullOrWhiteSpace(paymentRequest.AdministrativeServiceNotificationURL))
                    {
                        EserviceNotification statusNotification = new EserviceNotification(paymentRequest);

                        this.systemRepository.AddEntity<EserviceNotification>(statusNotification);
                        this.unitOfWork.Save();
                    }

                    return new object();
                }
                else
                {
                    throw new CustomServiceException("Unable to suspend payment request with status different from 'Pending'.");
                }
            }
            else
            {
                throw new CustomServiceException("Payment request not found.");
            }
        }

        [Route("eik")]
        [HttpPost]
        public object PostEikRequest(AuthRequestDO requestDO)
        {
            string eik;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                eik = dataJObject.GetValue("eik", StringComparison.OrdinalIgnoreCase).Value<string>();
            }

            catch(Exception ex)
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var eikrequestTuples = this.apiRepository.GetParsedRequestInfoByEik(eik);

            //var departments = eikrequestTuples.Select(d => new { Name = d.Item2.Name, UniqueIdentificationNumber = d.Item2.UniqueIdentificationNumber }).ToList();
            //var aisclients = eikrequestTuples.Select(a => new { EserviceClientId = a.Item2.EserviceClientId, AisName = a.Item2.AisName }).ToList();
            //var banks = eikrequestTuples.Select(b => new { AccountBank = b.Item2.AccountBank, BIC = b.Item2.AccountBIC, IBAN = b.Item2.AccountIBAN }).ToList();


            return new
            {
                Eikrequest = eikrequestTuples.Select(e => new { eik = e.Item1, UniqueIdentificationNumber = e.Item2 }).ToList()
            };
        }

        [Route("refid")]
        [HttpPost]
        public object PostRefidRequestJson(AuthRequestDO requestDO)
        {
            int refid;
            string clientId;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                refid = dataJObject.GetValue("refid", StringComparison.OrdinalIgnoreCase).Value<int>();
                clientId = requestDO.ClientId;
            }

            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var refidrequestTuple = this.apiRepository.GetParsedRequestInfoByRefid(refid, clientId);

            return new
            {
                RefidRequest = refidrequestTuple.Select(r => new { refid = r.Item1, result = r.Item2 }).ToList()
            };
        }

        [Route("pendingPaymentStatusJson")]
        [HttpPost]
        public object PostPendingPaymentStatusJson(AuthRequestDO requestDO)
        {
            string applicantUin;
            int pageNumber;
            int pageSize;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                applicantUin = dataJObject.GetValue("applicantUin", StringComparison.OrdinalIgnoreCase).Value<string>();
                pageNumber = dataJObject.GetValue("pageNumber", StringComparison.OrdinalIgnoreCase).Value<int>();
                pageSize = dataJObject.GetValue("pageSize", StringComparison.OrdinalIgnoreCase).Value<int>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var pendingRequestTuples = this.apiRepository.GetParsedPendingPaymentRequestInfosByApplicant(applicantUin, pageNumber, pageSize);

            return new
            {
                count = pendingRequestTuples.Item1,
                PaymentStatuses = pendingRequestTuples.Item2.Select(r => new {paymentrequestID = r.Item1, result = r.Item2 }).ToList()
            };
        }

        [Route("paymentStatusJson")]
        [HttpPost]
        public object PostPaymentStatusJson(AuthRequestDO requestDO)
        {
            string applicantUin;
            int pageNumber;
            int pageSize;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                applicantUin = dataJObject.GetValue("applicantUin", StringComparison.OrdinalIgnoreCase).Value<string>();
                pageNumber = dataJObject.GetValue("pageNumber", StringComparison.OrdinalIgnoreCase).Value<int>();
                pageSize = dataJObject.GetValue("pageSize", StringComparison.OrdinalIgnoreCase).Value<int>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var pendingRequestTuples = this.apiRepository.GetParsedPaymentRequestInfosByApplicant(applicantUin, pageNumber, pageSize);

            return new
            { 
                count = pendingRequestTuples.Item1,
                PaymentStatuses = pendingRequestTuples.Item2.Select(r => new { paymentrequestID = r.Item1, result = r.Item2 }).ToList()
            };
        }

        [Route("paymentStatusJsonGrouped")]
        [HttpPost]
        public object PostPaymentStatusJsonGrouped(AuthRequestDO requestDO)
        {
            string applicantUin;
            int pageNumber;
            int pageSize;
            PaymentRequestStatus[] paymentStatuses;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                applicantUin = dataJObject.GetValue("applicantUin", StringComparison.OrdinalIgnoreCase).Value<string>();
                paymentStatuses = dataJObject.GetValue("paymentStatuses", StringComparison.OrdinalIgnoreCase).Value<JArray>().ToObject<PaymentRequestStatus[]>();
                pageNumber = dataJObject.GetValue("pageNumber", StringComparison.OrdinalIgnoreCase).Value<int>();
                pageSize = dataJObject.GetValue("pageSize", StringComparison.OrdinalIgnoreCase).Value<int>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var pendingRequestTuples = this.apiRepository.GetParsedPaymentRequestInfosByApplicantGrouped(applicantUin, pageNumber, pageSize, paymentStatuses);

            return new
            {
                count = pendingRequestTuples.Item1,
                PaymentStatuses = pendingRequestTuples.Item2.Select(r => new { paymentrequestStatus = r.Item1, result = r.Item2 }).ToList()
            };
        }

        [Route("numberOfPaymentsByMonthJson")]
        [HttpPost]
        public object PostNumberOfPaymentsByMonthJson(AuthRequestDO requestDO)
        {
            string request_type;
            string month;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                request_type = dataJObject.GetValue("request_type", StringComparison.OrdinalIgnoreCase).Value<string>();
                month = dataJObject.GetValue("month", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var numberOfPayments = this.apiRepository.GetNumberOfPaymentsByMonth(request_type, month);

            return new
            {
                month = month,
                value = numberOfPayments
            };
        }

        [Route("numberOfPaymentsFromDateToDateJson")]
        [HttpPost]
        public object PostNumberOfPaymentsFromDateToDateJson(AuthRequestDO requestDO)
        {
            string request_type;
            string start_date;
            string end_date;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                request_type = dataJObject.GetValue("request_type", StringComparison.OrdinalIgnoreCase).Value<string>();
                start_date = dataJObject.GetValue("start_date", StringComparison.OrdinalIgnoreCase).Value<string>();
                end_date = dataJObject.GetValue("end_date", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var numberOfPayments = this.apiRepository.GetNumberOfPaymentsFromDateToDate(request_type, start_date, end_date);
            string fromdateToDate = $"{start_date} to {end_date}";

            return new
            {
                from = fromdateToDate,
                value = numberOfPayments
            };
        }

        [Route("numberOfVposePayCvposPaidRequestsJson")]
        [HttpPost]
        public object PostNumberOfVposePayCvposPaidRequestsJson(AuthRequestDO requestDO)
        {
            string start_date;
            string end_date;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                start_date = dataJObject.GetValue("start_date", StringComparison.OrdinalIgnoreCase).Value<string>();
                end_date = dataJObject.GetValue("end_date", StringComparison.OrdinalIgnoreCase).Value<string>();
            }

            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var numberOfPayments = this.apiRepository.GetNumberOfVposePayCvposPaidRequestsFromDateToDate(start_date, end_date);
            //string fromdateToDate = $"{start_date} to {end_date}";

            return numberOfPayments;
        }

        [Route("administrationsWithTheHighestNumberOfPaymentsJson")]
        [HttpPost]
        public object PostAdministrationsWithTheHighestNumberOfPaymentsJson(AuthRequestDO requestDO)
        {
            string start_date;
            string end_date;
            try
            {
                string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
                JObject dataJObject = JObject.Parse(dataJson);
                start_date = dataJObject.GetValue("start_date", StringComparison.OrdinalIgnoreCase).Value<string>();
                end_date = dataJObject.GetValue("end_date", StringComparison.OrdinalIgnoreCase).Value<string>();
            }
            catch
            {
                throw new CustomServiceException("An error occurred when attempting to parse data. Please ensure that POST data is in correct format.");
            }

            var numberOfPayments = this.apiRepository.GetAdministrationsWithTheHighestNumberOfPaymentsFromDateToDate(start_date, end_date);
            //string fromdateToDate = $"{start_date} to {end_date}";

            return numberOfPayments;
        }

        #region Private members

        private Tuple<R_10052.PaymentRequestUnacceptedReceipt, R_10055.PaymentRequestAcceptedReceipt> InternalPostPayment(
            AuthRequestDO requestDO,
            string requestXmlContent,
            string aisPaymentId,
            bool isExtendet = false,
            string eServiceClientId = null,
            int? payOrder = null)
        {
            List<string> schemaErrors = this.schemaValidator.ValidateXmlSchema(requestXmlContent);

            if (schemaErrors.Count > 0)
            {
                throw new CustomServiceException(BuildSchemaExceptionErrorMsg(schemaErrors));
            }

            R_10046.PaymentRequest requestRio = this.documentSerializer.XmlDeserializeFromString<R_10046.PaymentRequest>(requestXmlContent);

            PaymentRequest request = null;

            List<string> validationErrors = new List<string>();

            EserviceClient еserviceClient = this.systemRepository.GetEserviceClientByClientId(isExtendet ? eServiceClientId : requestDO.ClientId);
            if (еserviceClient != null)
            {
                validationErrors.AddRange(DocumentHelper.ValidatePaymentRequestData(apiRepository, requestRio, еserviceClient.EserviceClientId, aisPaymentId, this.paymentRequestRepository, out request));
            }
            else
            {
                validationErrors.Add($"Can not find EService client with EserviceClientId {(isExtendet ? eServiceClientId : requestDO.ClientId)}");
            }
            bool isNewRequest = request == null;

            bool isRequestValid = validationErrors.Count == 0;

            string unacceptedReceiptXml = null;
            string acceptedReceiptXml = null;
            string paymentRequestIdentifier = null;

            R_10055.PaymentRequestAcceptedReceipt acceptedReceiptRio = null;
            R_10052.PaymentRequestUnacceptedReceipt unacceptedReceiptRio = null;

            if (isRequestValid)
            {
                if (isNewRequest)
                    paymentRequestIdentifier = GetPaymentRequestIdentifier();
                else
                    paymentRequestIdentifier = request.PaymentRequestIdentifier;

                acceptedReceiptRio = new R_10055.PaymentRequestAcceptedReceipt();
                acceptedReceiptRio.PaymentRequestID = paymentRequestIdentifier;
                acceptedReceiptRio.PaymentRequestRegistrationTime = DateTime.Now;
                acceptedReceiptRio.DocumentTypeName = "Съобщение, че заявката за плащане се приема";
                acceptedReceiptRio.DocumentTypeURI = new R_0009_000003.DocumentTypeURI();
                acceptedReceiptRio.DocumentTypeURI.BatchNumber = "0010";
                acceptedReceiptRio.DocumentTypeURI.RegisterIndex = "000001";

                acceptedReceiptXml = this.documentSerializer.XmlSerializeToString(acceptedReceiptRio);
            }
            else
            {
                unacceptedReceiptRio = new R_10052.PaymentRequestUnacceptedReceipt();
                unacceptedReceiptRio.PaymentRequest = requestRio;
                unacceptedReceiptRio.PaymentRequestValidationTime = DateTime.Now;
                unacceptedReceiptRio.DocumentTypeName = "Съобщение, че заявката за плащане не се приема";
                unacceptedReceiptRio.DocumentTypeURI = new R_0009_000003.DocumentTypeURI();
                unacceptedReceiptRio.DocumentTypeURI.BatchNumber = "0010";
                unacceptedReceiptRio.DocumentTypeURI.RegisterIndex = "000001";
                unacceptedReceiptRio.PaymentRequestErrors = new R_10052.PaymentRequestErrors();
                unacceptedReceiptRio.PaymentRequestErrors.PaymentRequestErrorCollection = new R_10052.PaymentRequestErrorCollection();

                foreach (string error in validationErrors)
                {
                    unacceptedReceiptRio.PaymentRequestErrors.PaymentRequestErrorCollection.Add(new R_10036.PaymentRequestError()
                    {
                        TermName = error
                    });
                }

                unacceptedReceiptXml = this.documentSerializer.XmlSerializeToString(unacceptedReceiptRio);
            }

            User user = null;

            using (var transaction = this.unitOfWork.BeginTransaction())
            {
                PaymentRequestXml requestXml = new PaymentRequestXml();
                requestXml.RequestContent = requestXmlContent;
                requestXml.ResponseContent = isRequestValid ? acceptedReceiptXml : unacceptedReceiptXml;
                requestXml.IsRequestAccepted = isRequestValid;
                requestXml.CreateDate = DateTime.Now;

                if (isRequestValid)
                {
                    var uinData = DocumentHelper.GetUinTypeUinAndName(requestRio.ElectronicServiceRecipient);

                    if (isNewRequest)
                    {
                        request = new PaymentRequest();

                        SetPaymentRequestAttributes(request, requestRio, uinData);

                        request.PayOrder = payOrder;
                        request.PaymentRequestXml = requestXml;
                        request.EserviceClientId = еserviceClient.EserviceClientId;
                        request.InitiatorId = еserviceClient.EserviceClientId;
                        request.InitiatorEserviceClient = еserviceClient;

                        if (isExtendet) 
                        {
                            EserviceClient еserviceClientCheck = this.systemRepository.GetEserviceClientByClientId(requestDO.ClientId);
                            if (еserviceClientCheck == null) 
                            {
                                throw new CustomServiceException("Service Client can not be found");
                            }
                            request.InitiatorId = еserviceClientCheck.EserviceClientId;
                            request.InitiatorEserviceClient = еserviceClientCheck;
                        }
                        
                        request.Gid = Guid.NewGuid();
                        request.AisPaymentId = aisPaymentId;
                        request.PaymentRequestIdentifier = paymentRequestIdentifier;
                        request.PaymentRequestStatusId = PaymentRequestStatus.Pending;
                        request.PaymentRequestStatusChangeTime = DateTime.Now;
                        request.CreateDate = DateTime.Now;
                        request.PaymentRequestAccessCode = this.commonRepository.GeneratePaymentRequestAccessCode();
                        request.ObligationStatusId = Model.Enums.ObligationStatusEnum.Asked;

                        this.apiRepository.AddEntity<PaymentRequest>(request);
                    }
                    else
                    {
                        SetPaymentRequestAttributes(request, requestRio, uinData);

                        request.PaymentRequestXml = requestXml;
                        request.CreateDate = DateTime.Now;
                    }

                    //Create user if need
                    user = this.systemRepository.GetUserByUin(request.ApplicantUin);
                    if (user == null)
                    {
                        user = new User();
                        user.Egn = request.ApplicantUin;
                        user.RequestNotifications = false;
                        user.StatusNotifications = false;
                        user.StatusObligationNotifications = false;

                        this.systemRepository.AddEntity<User>(user);
                    }
                }
                else
                {
                    this.apiRepository.AddEntity<PaymentRequestXml>(requestXml);
                }

                this.unitOfWork.Save();

                transaction.Commit();
            }

            if (isNewRequest)
            {
                if (AppSettings.EPaymentsEventRegister_Enabled)
                {
                    EventRegisterNotification notification = new EventRegisterNotification(
                        isRequestValid ? request.PaymentRequestId : (int?)null,
                        isRequestValid ? EventRegisterNotificationType.PaymentRequestRegistered : EventRegisterNotificationType.PaymentRequestDenied,
                        isRequestValid ? String.Format("(еПлащане) Регистрирана заявка за плащане с №{0}", request.PaymentRequestIdentifier) : "(еПлащане) Oтказана е регистрация на заявка за плащане",
                        isRequestValid ? request.PaymentRequestIdentifier : null);

                    this.systemRepository.AddEntity<EventRegisterNotification>(notification);
                    this.unitOfWork.Save();
                }
            }

            if (isRequestValid &&
                isNewRequest &&
                user.RequestNotifications &&
                !String.IsNullOrWhiteSpace(user.Email))
            {
                NewPaymentRequestContextDO context = new NewPaymentRequestContextDO(
                    request.PaymentRequestIdentifier,
                    request.ServiceProviderName,
                    request.PaymentReason,
                    request.PaymentAmount,
                    request.PaymentRequestAccessCode,
                    request.CreateDate);

                Email email = new Email(context, user.Email);
                this.systemRepository.AddEntity<Email>(email);

                this.unitOfWork.Save();
            }
            if (isRequestValid &&
               isNewRequest &&
               user.StatusObligationNotifications &&
               !String.IsNullOrWhiteSpace(user.Email))
            {
                StatusChangedObligationContextDO contextObl = new StatusChangedObligationContextDO(
                    request.PaymentRequestIdentifier,
                    request.ServiceProviderName,
                    request.PaymentReason,
                    request.PaymentAmount,
                    request.ObligationStatusId.GetDescription(), 
                    request.CreateDate);

                Email email = new Email(contextObl, user.Email);
                this.systemRepository.AddEntity<Email>(email);

                this.unitOfWork.Save();
            }

            return new Tuple<R_10052.PaymentRequestUnacceptedReceipt, R_10055.PaymentRequestAcceptedReceipt>(
                unacceptedReceiptRio,
                acceptedReceiptRio);
        }

        private void SetPaymentRequestAttributes(PaymentRequest request, R_10046.PaymentRequest requestRio, Tuple<UinType, string, string> uinData)
        {
            request.IsPaymentMultiple = false;
            request.ServiceProviderName = requestRio.ElectronicServiceProviderBasicData?.EntityBasicData?.Name;
            request.ServiceProviderBank = requestRio.ElectronicServiceProviderBankAccount?.EntityBasicData?.Name;
            request.ServiceProviderBIC = requestRio.ElectronicServiceProviderBankAccount?.BIC;
            request.ServiceProviderIBAN = requestRio.ElectronicServiceProviderBankAccount?.IBAN?.Trim();
            request.Currency = requestRio.Currency;
            request.PaymentTypeCode = requestRio.PaymentCode;
            request.ObligationTypeId = requestRio.ObligationType != null ? int.Parse(requestRio.ObligationType) : 1;
            request.PaymentAmount = Convert.ToDecimal(requestRio.PaymentAmount);
            request.PaymentReason = requestRio.PaymentReason;
            request.ApplicantUinTypeId = uinData.Item1;
            request.ApplicantUin = uinData.Item2;
            request.ApplicantName = Regex.Replace(uinData.Item3, @"\s+", " ");
            request.PaymentReferenceType = requestRio.PaymentReference?.PaymentReferenceType?.Trim();
            request.PaymentReferenceNumber = requestRio.PaymentReference?.PaymentReferenceNumber?.Trim();
            request.PaymentReferenceDate = requestRio.PaymentReference.PaymentReferenceDate.Value;
            request.ExpirationDate = requestRio.PaymentRequestExpirationDate.Value;
            request.AdditionalInformation = requestRio.AdditionalInformationInPaymentRequest;
            request.AdministrativeServiceUri = requestRio.SUNAUServiceURI;
            request.AdministrativeServiceSupplierUri = requestRio.ElectronicAdministrativeServiceSupplierUriRA;
            request.AdministrativeServiceNotificationURL = requestRio.ElectronicAdministrativeServiceNotificationURL;
        }

        private R_10046.PaymentRequest CreatePaymentRequestRioFromJson(PaymentRequestJsonDO paymentRequestJson)
        {
            R_10046.PaymentRequest request = new R_10046.PaymentRequest();
            request.DocumentTypeURI = new R_0009_000003.DocumentTypeURI();
            request.DocumentTypeURI.RegisterIndex = "000001";
            request.DocumentTypeURI.BatchNumber = "0010";
            request.DocumentTypeName = "Заявка за плащане";
            request.PaymentCode = paymentRequestJson.PaymentTypeCode;
            request.ObligationType = paymentRequestJson.ObligationType;
            request.Currency = paymentRequestJson.Currency;
            request.PaymentAmount = paymentRequestJson.PaymentAmount.Value;
            request.ElectronicServiceProviderBasicData = new R_0009_000002.ElectronicServiceProviderBasicData();
            request.ElectronicServiceProviderBasicData.EntityBasicData = new R_0009_000013.EntityBasicData();
            request.ElectronicServiceProviderBasicData.EntityBasicData.Name = paymentRequestJson.ServiceProviderName;
            request.ElectronicServiceProviderBasicData.EntityBasicData.Identifier = null;
            request.ElectronicServiceProviderBasicData.ElectronicServiceProviderType = "0006-000031";
            request.ElectronicServiceProviderBankAccount = new R_10010.ElectronicServiceProviderBankAccount();
            request.ElectronicServiceProviderBankAccount.BIC = paymentRequestJson.ServiceProviderBIC;
            request.ElectronicServiceProviderBankAccount.IBAN = paymentRequestJson.ServiceProviderIBAN;
            request.ElectronicServiceProviderBankAccount.EntityBasicData = new R_0009_000013.EntityBasicData();
            request.ElectronicServiceProviderBankAccount.EntityBasicData.Name = paymentRequestJson.ServiceProviderBank;

            request.ElectronicServiceRecipient = new R_0009_000015.ElectronicServiceRecipient();
            if (paymentRequestJson.ApplicantUinTypeId == (int)UinType.Egn)
            {
                request.ElectronicServiceRecipient.Person = new R_0009_000008.PersonBasicData();
                request.ElectronicServiceRecipient.Person.Identifier = new R_0009_000006.PersonIdentifier();
                request.ElectronicServiceRecipient.Person.Identifier.EGN = paymentRequestJson.ApplicantUin;
                request.ElectronicServiceRecipient.Person.Names = new R_0009_000005.PersonNames();

                Tuple<string, string, string> tuple = Formatter.SplitNames(paymentRequestJson.ApplicantName);

                request.ElectronicServiceRecipient.Person.Names.First = (!string.IsNullOrWhiteSpace(tuple.Item1) ? tuple.Item1 : null);
                request.ElectronicServiceRecipient.Person.Names.Middle = (!string.IsNullOrWhiteSpace(tuple.Item2) ? tuple.Item2 : null);
                request.ElectronicServiceRecipient.Person.Names.Last = (!string.IsNullOrWhiteSpace(tuple.Item3) ? tuple.Item3 : null);
            }
            else if (paymentRequestJson.ApplicantUinTypeId == (int)UinType.Lnch)
            {
                request.ElectronicServiceRecipient.Person = new R_0009_000008.PersonBasicData();
                request.ElectronicServiceRecipient.Person.Identifier = new R_0009_000006.PersonIdentifier();
                request.ElectronicServiceRecipient.Person.Identifier.LNCh = paymentRequestJson.ApplicantUin;
                request.ElectronicServiceRecipient.Person.Names = new R_0009_000005.PersonNames();

                Tuple<string, string, string> tuple = Formatter.SplitNames(paymentRequestJson.ApplicantName);

                request.ElectronicServiceRecipient.Person.Names.First = (!string.IsNullOrWhiteSpace(tuple.Item1) ? tuple.Item1 : null);
                request.ElectronicServiceRecipient.Person.Names.Middle = (!string.IsNullOrWhiteSpace(tuple.Item2) ? tuple.Item2 : null);
                request.ElectronicServiceRecipient.Person.Names.Last = (!string.IsNullOrWhiteSpace(tuple.Item3) ? tuple.Item3 : null);
            }
            else if (paymentRequestJson.ApplicantUinTypeId == (int)UinType.Bulstat)
            {
                request.ElectronicServiceRecipient.Entity = new R_0009_000013.EntityBasicData();
                request.ElectronicServiceRecipient.Entity.Identifier = paymentRequestJson.ApplicantUin;
                request.ElectronicServiceRecipient.Entity.Name = paymentRequestJson.ApplicantName;
            }

            request.PaymentReason = paymentRequestJson.PaymentReason;
            request.PaymentReference = new R_10018.PaymentReference();
            request.PaymentReference.PaymentReferenceType = paymentRequestJson.PaymentReferenceType;
            request.PaymentReference.PaymentReferenceNumber = paymentRequestJson.PaymentReferenceNumber;
            request.PaymentReference.PaymentReferenceDate = paymentRequestJson.PaymentReferenceDate;
            request.PaymentPeriod = new R_10024.PaymentPeriod();
            request.PaymentPeriod.PaymentPeriodFromDate = paymentRequestJson.PaymentReferenceDate;
            request.PaymentPeriod.PaymentPeriodToDate = paymentRequestJson.PaymentReferenceDate;
            request.PaymentRequestExpirationDate = paymentRequestJson.ExpirationDate;
            request.AdditionalInformationInPaymentRequest = paymentRequestJson.AdditionalInformation;

            request.SUNAUServiceURI = paymentRequestJson.AdministrativeServiceUri;
            request.ElectronicAdministrativeServiceSupplierUriRA = paymentRequestJson.AdministrativeServiceSupplierUri;
            request.ElectronicAdministrativeServiceNotificationURL = paymentRequestJson.AdministrativeServiceNotificationURL;

            return request;
        }
        private R_10046.PaymentRequest CreatePaymentRequestRioFromJsonExtendet(PaymentRequestJsonExtendedDO paymentRequestJson, string eServiceClientId)
        {
            EserviceClient еserviceClient = this.systemRepository.GetEserviceClientByClientId(eServiceClientId);

            R_10046.PaymentRequest request = new R_10046.PaymentRequest();
            request.DocumentTypeURI = new R_0009_000003.DocumentTypeURI();
            request.DocumentTypeURI.RegisterIndex = "000001";
            request.DocumentTypeURI.BatchNumber = "0010";
            request.DocumentTypeName = "Заявка за плащане";
            request.PaymentCode = paymentRequestJson.PaymentTypeCode;
            request.ObligationType = paymentRequestJson.ObligationType;
            request.Currency = paymentRequestJson.Currency;
            request.PaymentAmount = paymentRequestJson.PaymentAmount.Value;
            request.ElectronicServiceProviderBasicData = new R_0009_000002.ElectronicServiceProviderBasicData();
            request.ElectronicServiceProviderBasicData.EntityBasicData = new R_0009_000013.EntityBasicData();
            request.ElectronicServiceProviderBasicData.EntityBasicData.Name = еserviceClient?.ServiceName;
            request.ElectronicServiceProviderBasicData.EntityBasicData.Identifier = null;
            request.ElectronicServiceProviderBasicData.ElectronicServiceProviderType = "0006-000031";
            request.ElectronicServiceProviderBankAccount = new R_10010.ElectronicServiceProviderBankAccount();
            request.ElectronicServiceProviderBankAccount.BIC = еserviceClient?.AccountBIC;
            request.ElectronicServiceProviderBankAccount.IBAN = еserviceClient?.AccountIBAN;
            request.ElectronicServiceProviderBankAccount.EntityBasicData = new R_0009_000013.EntityBasicData();
            request.ElectronicServiceProviderBankAccount.EntityBasicData.Name = еserviceClient?.AccountBank;

            request.ElectronicServiceRecipient = new R_0009_000015.ElectronicServiceRecipient();
            if (paymentRequestJson.ApplicantUinTypeId == (int)UinType.Egn)
            {
                request.ElectronicServiceRecipient.Person = new R_0009_000008.PersonBasicData();
                request.ElectronicServiceRecipient.Person.Identifier = new R_0009_000006.PersonIdentifier();
                request.ElectronicServiceRecipient.Person.Identifier.EGN = paymentRequestJson.ApplicantUin;
                request.ElectronicServiceRecipient.Person.Names = new R_0009_000005.PersonNames();

                Tuple<string, string, string> tuple = Formatter.SplitNames(paymentRequestJson.ApplicantName);

                request.ElectronicServiceRecipient.Person.Names.First = (!string.IsNullOrWhiteSpace(tuple.Item1) ? tuple.Item1 : null);
                request.ElectronicServiceRecipient.Person.Names.Middle = (!string.IsNullOrWhiteSpace(tuple.Item2) ? tuple.Item2 : null);
                request.ElectronicServiceRecipient.Person.Names.Last = (!string.IsNullOrWhiteSpace(tuple.Item3) ? tuple.Item3 : null);
            }
            else if (paymentRequestJson.ApplicantUinTypeId == (int)UinType.Lnch)
            {
                request.ElectronicServiceRecipient.Person = new R_0009_000008.PersonBasicData();
                request.ElectronicServiceRecipient.Person.Identifier = new R_0009_000006.PersonIdentifier();
                request.ElectronicServiceRecipient.Person.Identifier.LNCh = paymentRequestJson.ApplicantUin;
                request.ElectronicServiceRecipient.Person.Names = new R_0009_000005.PersonNames();

                Tuple<string, string, string> tuple = Formatter.SplitNames(paymentRequestJson.ApplicantName);

                request.ElectronicServiceRecipient.Person.Names.First = (!string.IsNullOrWhiteSpace(tuple.Item1) ? tuple.Item1 : null);
                request.ElectronicServiceRecipient.Person.Names.Middle = (!string.IsNullOrWhiteSpace(tuple.Item2) ? tuple.Item2 : null);
                request.ElectronicServiceRecipient.Person.Names.Last = (!string.IsNullOrWhiteSpace(tuple.Item3) ? tuple.Item3 : null);
            }
            else if (paymentRequestJson.ApplicantUinTypeId == (int)UinType.Bulstat)
            {
                request.ElectronicServiceRecipient.Entity = new R_0009_000013.EntityBasicData();
                request.ElectronicServiceRecipient.Entity.Identifier = paymentRequestJson.ApplicantUin;
                request.ElectronicServiceRecipient.Entity.Name = paymentRequestJson.ApplicantName;
            }

            request.PaymentReason = paymentRequestJson.PaymentReason;
            request.PaymentReference = new R_10018.PaymentReference();
            request.PaymentReference.PaymentReferenceType = paymentRequestJson.PaymentReferenceType;
            request.PaymentReference.PaymentReferenceNumber = paymentRequestJson.PaymentReferenceNumber;
            request.PaymentReference.PaymentReferenceDate = paymentRequestJson.PaymentReferenceDate;
            request.PaymentPeriod = new R_10024.PaymentPeriod();
            request.PaymentPeriod.PaymentPeriodFromDate = paymentRequestJson.PaymentReferenceDate;
            request.PaymentPeriod.PaymentPeriodToDate = paymentRequestJson.PaymentReferenceDate;
            request.PaymentRequestExpirationDate = paymentRequestJson.ExpirationDate;
            request.AdditionalInformationInPaymentRequest = paymentRequestJson.AdditionalInformation;
            request.SUNAUServiceURI = paymentRequestJson.AdministrativeServiceUri;
            request.ElectronicAdministrativeServiceSupplierUriRA = paymentRequestJson.AdministrativeServiceSupplierUri;
            request.ElectronicAdministrativeServiceNotificationURL = paymentRequestJson.AdministrativeServiceNotificationURL;

            return request;
        }

        private UnacceptedReceiptJsonDO CreateUnacceptedReceiptJsonDO(R_10052.PaymentRequestUnacceptedReceipt unacceptedReceipt)
        {
            UnacceptedReceiptJsonDO returnValue = new UnacceptedReceiptJsonDO();

            returnValue.ValidationTime = unacceptedReceipt.PaymentRequestValidationTime.Value;
            returnValue.Errors = new List<string>();

            foreach (var requestError in unacceptedReceipt.PaymentRequestErrors.PaymentRequestErrorCollection)
            {
                returnValue.Errors.Add(requestError.TermName);
            }

            return returnValue;
        }

        private AcceptedReceiptJsonDO CreateAcceptedReceiptJsonDO(R_10055.PaymentRequestAcceptedReceipt acceptedReceipt)
        {

            AcceptedReceiptJsonDO returnValue = new AcceptedReceiptJsonDO();

            returnValue.Id = acceptedReceipt.PaymentRequestID;
            returnValue.RegistrationTime = acceptedReceipt.PaymentRequestRegistrationTime.Value;

            return returnValue;
        }

        private PaymentRequestJsonDO CreatePaymentRequestJsonDO(PaymentRequest paymentRequest)
        {
            PaymentRequestJsonDO paymentRequestJson = new PaymentRequestJsonDO();

            paymentRequestJson.ServiceProviderName = paymentRequest.ServiceProviderName;
            paymentRequestJson.ServiceProviderBank = paymentRequest.ServiceProviderBank;
            paymentRequestJson.ServiceProviderBIC = paymentRequest.ServiceProviderBIC;
            paymentRequestJson.ServiceProviderIBAN = paymentRequest.ServiceProviderIBAN;
            paymentRequestJson.Currency = paymentRequest.Currency;
            paymentRequestJson.PaymentTypeCode = paymentRequest.PaymentTypeCode;
            paymentRequestJson.PaymentAmount = (double)paymentRequest.PaymentAmount;
            paymentRequestJson.PaymentReason = paymentRequest.PaymentReason;
            paymentRequestJson.ApplicantUinTypeId = (int)paymentRequest.ApplicantUinTypeId;
            paymentRequestJson.ApplicantUin = paymentRequest.ApplicantUin;
            paymentRequestJson.ApplicantName = paymentRequest.ApplicantName;
            paymentRequestJson.PaymentReferenceType = paymentRequest.PaymentReferenceType;
            paymentRequestJson.PaymentReferenceNumber = paymentRequest.PaymentReferenceNumber;
            paymentRequestJson.PaymentReferenceDate = paymentRequest.PaymentReferenceDate;
            paymentRequestJson.ExpirationDate = paymentRequest.ExpirationDate;
            paymentRequestJson.AdditionalInformation = paymentRequest.AdditionalInformation;
            paymentRequestJson.AdministrativeServiceUri = paymentRequest.AdministrativeServiceUri;
            paymentRequestJson.AdministrativeServiceSupplierUri = paymentRequest.AdministrativeServiceSupplierUri;
            paymentRequestJson.AdministrativeServiceNotificationURL = paymentRequest.AdministrativeServiceNotificationURL;

            return paymentRequestJson;
        }

        private Authorizer GetEserviceAuthorizer(string clientId)
        {
            return new Authorizer(apiRepository, clientId);
        }

        private string GetPaymentRequestIdentifier()
        {
            DateTime nowDate = DateTime.Now;
            int? counter = null;

            using (var transaction = this.unitOfWork.BeginTransaction())
            {
                counter = this.systemRepository.GetPaymentRequestCounter(nowDate);

                transaction.Commit();
            }

            return String.Format("{0}{1}", Formatter.DateToShortFormat(nowDate), counter.Value);
        }

        private string BuildSchemaExceptionErrorMsg(List<string> errors)
        {
            string errorMsg = String.Empty;

            if (errors != null && errors.Count > 0)
            {
                StringBuilder builder = new StringBuilder();

                builder.AppendLine("Xsd schema validation errors occurred:");

                for (int i = 0; i < errors.Count; i++)
                {
                    builder.AppendLine((i + 1).ToString() + ". " + errors[i]);
                }

                errorMsg = builder.ToString();
            }

            return errorMsg;
        }

        #endregion
    }
}
