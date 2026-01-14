using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using System;
using System.Web.Mvc;
using EPayments.Model.Models;
using EPayments.Web.Common;
using EPayments.Data.ViewObjects.Web;
using System.Web.Script.Serialization;
using EPayments.Model.Enums;
using EPayments.Web.VposHelpers.Dsk;
using EPayments.Web.Auth;
using EPayments.Common.DataObjects;
using System.Text;
using Newtonsoft.Json.Linq;
using EPayments.Web.VposHelpers;
using EPayments.Common.Helpers;
using System.Web;
using EPayments.Web.Models.Shared;
using EPayments.Common;
using EPayments.Model.DataObjects.EmailTemplateContext;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Linq;
using System.Net.Security;
using EPayments.Web.VposHelpers.Borica;
using EPayments.Web.DataObjects;
using EPayments.Web.VposHelpers.FiBank;
using System.Data.Entity;
using EPayments.Common.VposHelpers;
using EPayments.Web.VposHelpers.Epay;
using System.Threading;
using System.Collections.Generic;
using EPayments.Data.ViewObjects.Web.APGModels;
using System.Security.Cryptography;
using System.Globalization;
using EPayments.Data.ViewObjects.Web.APGModels.Requests;
using EPayments.Common.BoricaHelpers;
using System.IO;
using log4net.Core;
using log4net;

namespace EPayments.Web.Controllers
{
    public partial class VposController : BaseController
    {
        private IUnitOfWork unitOfWork;
        private IWebRepository webRepository;
        private ISystemRepository systemRepository;
        private ILog logger;

        public VposController(IUnitOfWork unitOfWork, IWebRepository webRepository, ISystemRepository systemRepository)
        {
            this.unitOfWork = unitOfWork;
            this.webRepository = webRepository;
            this.systemRepository = systemRepository;
            this.logger = LogManager.GetLogger(typeof(VposController));
        }

        //External payment methods

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult TestEserviceClientVpos(AuthRequestDO requestDO)
        {
            try
            {
                EserviceClient eserviceClient =
                    this.unitOfWork.DbContext.Set<EserviceClient>()
                    .Include(e => e.VposClient)
                    .Where(e => e.ClientId == requestDO.ClientId)
                    .SingleOrDefault();

                if (eserviceClient.VposClientId.HasValue)
                {
                    if (eserviceClient.VposClientId.Value == (int)Vpos.DskEcomm)
                    {
                        VposRequestDataVO vposRequestDataVO = new VposRequestDataVO()
                        {
                            VposPaymentRequestUri = eserviceClient.VposClient.PaymentRequestUrl,
                            DskVposMerchantId = eserviceClient.DskVposMerchantId,
                            DskVposMerchantPassword = eserviceClient.DskVposMerchantPassword,
                            PaymentRequestIdentifier = "test",
                            PaymentAmount = 1.00M,
                            PaymentReason = "Test VPOS"
                        };

                        DskEcommVposRegisteredOrder registeredOrder = DskEcommVposService.RegisterOrder(
                            vposRequestDataVO.VposPaymentRequestUri,
                            vposRequestDataVO.DskVposMerchantId,
                            vposRequestDataVO.DskVposMerchantPassword,
                            vposRequestDataVO.PaymentRequestIdentifier,
                            vposRequestDataVO.PaymentAmount,
                            vposRequestDataVO.PaymentReason);

                        if (String.IsNullOrWhiteSpace(registeredOrder.OrderId) || 
                            (registeredOrder.ErrorCode != null && !registeredOrder.ErrorCode.Equals("0")) || 
                            registeredOrder.ErrorMessage != null)
                        {
                            throw new Exception($"Failed generating DskEcomm transaction id. ErrorCode: {registeredOrder.ErrorCode ?? string.Empty}; ErrorMessage: {registeredOrder.ErrorMessage ?? string.Empty}");
                        }

                        return Redirect(registeredOrder.FormUrl);
                    }
                    else if (eserviceClient.VposClientId.Value == (int)Vpos.Borica)
                    {
                        int boricaRequestIdentifier;

                        using (var transaction = this.unitOfWork.BeginTransaction())
                        {
                            boricaRequestIdentifier = this.systemRepository.GetPaymentRequestTotal();
                            transaction.Commit();
                        }

                        using (X509Certificate2 boricaCertificate = ReadBoricaCertificate(eserviceClient))
                        {
                            VposRequestDataVO vposRequestDataVO = new VposRequestDataVO()
                            {
                                PaymentAmount = 0.01M,
                                BoricaVposTerminalId = eserviceClient.BoricaVposTerminalId,
                                PaymentRequestIdentifier = boricaRequestIdentifier.ToString("000000"),
                                PaymentReason = "test",
                                BoricaVposBOReqSignCertFileName = eserviceClient.BoricaVposRequestSignCertFileName,
                                BoricaVposBOReqSignCertPassword = eserviceClient.BoricaVposRequestSignCertPassword,
                                VposPaymentRequestUri = eserviceClient.VposClient.PaymentRequestUrl
                            };

                            APGWPaymentRequestDataVO apgwRequest = BoricaVposHelper.CreateBoricaRequestModel(
                                vposRequestDataVO,
                                boricaRequestIdentifier,
                                eserviceClient,
                                boricaCertificate);

                            apgwRequest.P_Sign = this.CreatePSign(apgwRequest.GetPSignData(), boricaCertificate);

                            return this.View(MVC.Shared.Views._AutoPostBoricaVpos, apgwRequest);
                        }
                    }
                    if (eserviceClient.VposClientId.Value == (int)Vpos.FiBank)
                    {
                        VposRequestDataVO vposRequestDataVO = new VposRequestDataVO()
                        {
                            PaymentAmount = 1.00M,
                            PaymentRequestIdentifier = "test",
                            FiBankVposAccessKeystoreFilename = eserviceClient.FiBankVposAccessKeystoreFilename,
                            FiBankVposAccessKeystorePassword = eserviceClient.FiBankVposAccessKeystorePassword
                        };

                        string testIP = AppSettings.EPaymentsWeb_FiBankVposTestIp;

                        string transactionId = FiBankVposHelper.GenerateTransactionId(vposRequestDataVO, testIP);

                        if (String.IsNullOrWhiteSpace(transactionId))
                        {
                            throw new Exception("Failed generating Fibank transaction id.");
                        }

                        AutoPostVM model = FiBankVposHelper.CreateAutoPostVM(transactionId);

                        return View(MVC.Shared.Views._AutoPost, model);
                    }
                    else
                    {
                        throw new Exception("Invalid VPOS client id");
                    }
                }
                else
                {
                    return Content("No VPOS integrated with this AIS client.");
                }
            }
            catch (Exception ex)
            {
                return Content(Formatter.ExceptionToDetailedInfo(ex));
            }
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult Payment(AuthRequestDO requestDO)
        {
            ExternalPaymentRequestDO extPaymentRequestDO = ValidateAndGetExternalPaymentRequestDO(requestDO);

            if (!extPaymentRequestDO.IsRequestValid)
            {
                return AutoPostEservicePayment(
                    extPaymentRequestDO.EserviceClient.ClientId,
                    extPaymentRequestDO.EserviceClient.SecretKey,
                    extPaymentRequestDO.OkUrl,
                    extPaymentRequestDO.PaymentRequest.PaymentRequestIdentifier,
                    null,
                    extPaymentRequestDO.ErrorStatus.Value,
                    null);
            }

            if (extPaymentRequestDO.EserviceClient.VposClientId.Value == (int)Vpos.Dsk || extPaymentRequestDO.EserviceClient.VposClientId.Value == (int)Vpos.DskEcomm)
            {
                TaxAgreementDskVM model = new TaxAgreementDskVM();
                model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(extPaymentRequestDO.PaymentRequest.PaymentAmount,extPaymentRequestDO.PaymentRequest.CreateDate);
                model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(extPaymentRequestDO.PaymentRequest.PaymentAmount,extPaymentRequestDO.PaymentRequest.CreateDate);
                model.IsInternalPayment = false;
                model.ExternalRequestDO = requestDO;

                return View(MVC.Shared.Views._TaxAgreementDsk, model);
            }
            else if (extPaymentRequestDO.EserviceClient.VposClientId.Value == (int)Vpos.Borica)
            {
                TaxAgreementBoricaVM model = new TaxAgreementBoricaVM();
                model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(extPaymentRequestDO.PaymentRequest.PaymentAmount, extPaymentRequestDO.PaymentRequest.CreateDate);
                model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(extPaymentRequestDO.PaymentRequest.PaymentAmount, extPaymentRequestDO.PaymentRequest.CreateDate);
                model.IsInternalPayment = false;
                model.ExternalRequestDO = requestDO;

                return View(MVC.Shared.Views._TaxAgreementBorica, model);
            }
            else if (extPaymentRequestDO.EserviceClient.VposClientId.Value == (int)Vpos.FiBank)
            {
                TaxAgreementFiBankVM model = new TaxAgreementFiBankVM();
                model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(extPaymentRequestDO.PaymentRequest.PaymentAmount, extPaymentRequestDO.PaymentRequest.CreateDate);
                model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(extPaymentRequestDO.PaymentRequest.PaymentAmount, extPaymentRequestDO.PaymentRequest.CreateDate);
                model.IsInternalPayment = false;
                model.ExternalRequestDO = requestDO;

                return View(MVC.Shared.Views._TaxAgreementFiBank, model);
            }
            else
            {
                throw new HttpException(400, "Invalid VposClientId.");
            }
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult PaymentViaEpay(AuthRequestDO requestDO)
        {
            ExternalPaymentRequestDO extPaymentRequestDO = ValidateAndGetExternalPaymentRequestDO(requestDO);

            if (!extPaymentRequestDO.IsRequestValid)
            {
                return AutoPostEservicePayment(
                    extPaymentRequestDO.EserviceClient.ClientId,
                    extPaymentRequestDO.EserviceClient.SecretKey,
                    extPaymentRequestDO.OkUrl,
                    extPaymentRequestDO.PaymentRequest.PaymentRequestIdentifier,
                    null,
                    extPaymentRequestDO.ErrorStatus.Value,
                    null);
            }

            if (extPaymentRequestDO.EserviceClient.IsEpayVposEnabled)
            {
                TaxAgreementEpayVM model = new TaxAgreementEpayVM();
                model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(extPaymentRequestDO.PaymentRequest.PaymentAmount, extPaymentRequestDO.PaymentRequest.CreateDate);
                model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(extPaymentRequestDO.PaymentRequest.PaymentAmount, extPaymentRequestDO.PaymentRequest.CreateDate);
                model.IsInternalPayment = false;
                model.ExternalRequestDO = requestDO;

                return View(MVC.Shared.Views._TaxAgreementEpay, model);
            }
            else
            {
                throw new HttpException(400, "Invalid VposClientId.");
            }
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult PaymentConfirmed(AuthRequestDO requestDO, bool? isEpayPayment)
        {
            ExternalPaymentRequestDO extPaymentRequestDO = ValidateAndGetExternalPaymentRequestDO(requestDO);

            if (!extPaymentRequestDO.IsRequestValid)
            {
                return AutoPostEservicePayment(
                    extPaymentRequestDO.EserviceClient.ClientId,
                    extPaymentRequestDO.EserviceClient.SecretKey,
                    extPaymentRequestDO.OkUrl,
                    extPaymentRequestDO.PaymentRequest.PaymentRequestIdentifier,
                    null,
                    extPaymentRequestDO.ErrorStatus.Value,
                    null);
            }

            VposRedirect vposRedirect = new VposRedirect();
            vposRedirect.Gid = Guid.NewGuid();
            vposRedirect.PaymentRequestIdentifier = extPaymentRequestDO.PaymentRequest.PaymentRequestIdentifier;
            vposRedirect.OkUrl = extPaymentRequestDO.OkUrl;
            vposRedirect.CancelUrl = extPaymentRequestDO.CancelUrl;
            vposRedirect.IP = this.Request.UserHostAddress;
            vposRedirect.LogDate = DateTime.Now;

            this.systemRepository.AddEntity<VposRedirect>(vposRedirect);
            this.unitOfWork.Save();

            VposRequestDataVO dataVO = this.webRepository.GetVposRequestDataByGidAndUin(extPaymentRequestDO.PaymentRequest.Gid, extPaymentRequestDO.PaymentRequest.ApplicantUin);

            if (isEpayPayment == false && (extPaymentRequestDO.EserviceClient.VposClientId.Value == (int)Vpos.Dsk || extPaymentRequestDO.EserviceClient.VposClientId.Value == (int)Vpos.DskEcomm))
            {
                if (dataVO.VposClientId.Value == (int)Vpos.DskEcomm)
                {
                    return AutoPostDskEcomm(dataVO, vposRedirect.Gid, vposRedirect.VposRedirectId, vposRedirect.IP);
                }
                else
                {
                    return AutoPostDsk(dataVO, vposRedirect.Gid);
                }
            }
            else if (isEpayPayment == false && extPaymentRequestDO.EserviceClient.VposClientId.Value == (int)Vpos.Borica)
            {
                return AutoRedirectBorica(dataVO, vposRedirect.VposRedirectId);
            }
            else if (isEpayPayment == false && extPaymentRequestDO.EserviceClient.VposClientId.Value == (int)Vpos.FiBank)
            {
                return AutoPostFiBank(dataVO, vposRedirect.Gid, vposRedirect.VposRedirectId, vposRedirect.IP);
            }
            else if (isEpayPayment == true && extPaymentRequestDO.EserviceClient.IsEpayVposEnabled)
            {
                return AutoPostEpay(dataVO, vposRedirect.Gid, vposRedirect.VposRedirectId, vposRedirect.IP);
            }
            else
            {
                throw new HttpException(400, "Invalid VposClientId or POST parameters.");
            }
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult PaymentCanceled(AuthRequestDO requestDO)
        {
            ExternalPaymentRequestDO extPaymentRequestDO = ValidateAndGetExternalPaymentRequestDO(requestDO);

            return AutoPostEservicePayment(
                    extPaymentRequestDO.EserviceClient.ClientId,
                    extPaymentRequestDO.EserviceClient.SecretKey,
                    extPaymentRequestDO.CancelUrl,
                    extPaymentRequestDO.PaymentRequest.PaymentRequestIdentifier,
                    null,
                    EserviceVposResultStatus.CanceledByUser,
                    null);
        }

        [NonAction]
        private ExternalPaymentRequestDO ValidateAndGetExternalPaymentRequestDO(AuthRequestDO requestDO)
        {
            string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
            JObject dataJObject = JObject.Parse(dataJson);

            string id = dataJObject.GetValue("id", StringComparison.OrdinalIgnoreCase).Value<string>();
            string okUrl = dataJObject.GetValue("okUrl", StringComparison.OrdinalIgnoreCase).Value<string>();
            string cancelUrl = dataJObject.GetValue("cancelUrl", StringComparison.OrdinalIgnoreCase).Value<string>();

            if (String.IsNullOrWhiteSpace(id) || String.IsNullOrWhiteSpace(okUrl) || String.IsNullOrWhiteSpace(cancelUrl))
            {
                throw new HttpException(400, "Invalid post data.");
            }

            EserviceClient eserviceClient =
                this.unitOfWork.DbContext.Set<EserviceClient>()
                .Include(e => e.VposClient)
                .SingleOrDefault(e => e.ClientId == requestDO.ClientId);

            PaymentRequest request = this.webRepository.GetPaymentRequestByIdentifier(id);

            bool isRequestValid = false;
            EserviceVposResultStatus? vposResultStatus = null;

            if (request == null || request.EserviceClientId != eserviceClient.EserviceClientId)
            {
                throw new HttpException(400, "Payment request is invalid.");
            }
            else if (!eserviceClient.VposClientId.HasValue)
            {
                throw new HttpException(400, "No VirtualPOS is associated with eservice client.");
            }
            else if (request.PaymentRequestStatusId != PaymentRequestStatus.Pending || request.ExpirationDate <= DateTime.Now)
            {
                switch (request.PaymentRequestStatusId)
                {
                    case PaymentRequestStatus.Authorized:
                    case PaymentRequestStatus.Ordered:
                    case PaymentRequestStatus.Paid:
                        vposResultStatus = EserviceVposResultStatus.RequestIsPaid;
                        break;
                    case PaymentRequestStatus.Pending:
                    case PaymentRequestStatus.Expired:
                        vposResultStatus = EserviceVposResultStatus.RequestIsExpired;
                        break;
                    case PaymentRequestStatus.Canceled:
                    case PaymentRequestStatus.Suspended:
                        vposResultStatus = EserviceVposResultStatus.RequestIsCanceled;
                        break;
                    default:
                        throw new ArgumentException();
                }
            }
            else
            {
                isRequestValid = true;
            }

            return new ExternalPaymentRequestDO()
            {
                IsRequestValid = isRequestValid,
                ErrorStatus = vposResultStatus,
                OkUrl = okUrl,
                CancelUrl = cancelUrl,
                EserviceClient = eserviceClient,
                PaymentRequest = request
            };
        }

        //DSK methods

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult DskPayment(Guid id)
        {
            VposRequestDataVO dataVO;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                dataVO = this.webRepository.GetVposRequestDataByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                dataVO = this.webRepository.GetVposRequestDataByAccessCode(this.CurrentUser.AccessCode);
            }

            if (dataVO.VposClientId.Value == (int)Vpos.DskEcomm)
            {
                return AutoPostDskEcomm(dataVO, null, null, this.Request.UserHostAddress);
            }
            else
            {
                return AutoPostDsk(dataVO, null);
            }
        }

        [HttpGet]
        public virtual ActionResult DskSuccessfulPayment(Guid gid, Guid? vposRedirectGid, DskVposResultDO resultDO)
        {
            EserviceClient eserviceClient = this.systemRepository.GetEserviceClientByGid(gid);

            bool isRequestValid = DskVposHelper.IsRequestValid(eserviceClient, resultDO);

            if (!isRequestValid)
            {
                throw new Exception("Vpos request is invalid.");
            }

            string paymentRquestIdentifier = resultDO.Track_id.Substring(0, resultDO.Track_id.IndexOf('_'));

            PaymentRequest paymentRequest = this.webRepository.GetPaymentRequestByIdentifier(paymentRquestIdentifier);

            if (paymentRequest == null)
            {
                throw new Exception("Vpos request is invalid. PaymentRequest cannot be found.");
            }

            VposRedirect vposRedirect = null;
            if (vposRedirectGid.HasValue)
            {
                vposRedirect = this.systemRepository.GetVposRedirectByGid(vposRedirectGid.Value);
            }

            EserviceVposResultStatus resultStatus = resultDO.Response_code == "00" && resultDO.Status == "3" ? EserviceVposResultStatus.Success : EserviceVposResultStatus.Failure;

            return ProcessVposResultRequest(resultStatus, paymentRequest, eserviceClient, new JavaScriptSerializer().Serialize(resultDO), resultDO.Auth_code, vposRedirect);
        }

        [HttpGet]
        public virtual ActionResult DskFailedPayment(Guid gid, Guid? vposRedirectGid, DskVposResultDO resultDO)
        {
            EserviceClient eserviceClient = this.systemRepository.GetEserviceClientByGid(gid);

            bool isRequestValid = DskVposHelper.IsRequestValid(eserviceClient, resultDO);

            if (!isRequestValid)
            {
                throw new Exception("Vpos request is invalid.");
            }

            string paymentRquestIdentifier = resultDO.Track_id.Substring(0, resultDO.Track_id.IndexOf('_'));

            PaymentRequest paymentRequest = this.webRepository.GetPaymentRequestByIdentifier(paymentRquestIdentifier);

            if (paymentRequest == null)
            {
                throw new Exception("Vpos request is invalid. PaymentRequest cannot be found.");
            }

            VposRedirect vposRedirect = null;
            if (vposRedirectGid.HasValue)
            {
                vposRedirect = this.systemRepository.GetVposRedirectByGid(vposRedirectGid.Value);
            }

            EserviceVposResultStatus resultStatus = resultDO.Status == "8" ? EserviceVposResultStatus.CanceledByUser : EserviceVposResultStatus.Failure;

            return ProcessVposResultRequest(resultStatus, paymentRequest, eserviceClient, new JavaScriptSerializer().Serialize(resultDO), resultDO.Auth_code, vposRedirect);
        }

        [HttpGet]
        public virtual ActionResult SimulateVposPayment(Guid gid, Guid? vposRedirectGid, string paymentRequestIdentifier, bool success)
        {
            if (!AppSettings.EPaymentsWeb_SimulateVposPayment)
            {
                throw new Exception("Invalid operation");
            }

            DskVposResultDO resultDO = new DskVposResultDO();
            resultDO.Track_id = paymentRequestIdentifier;

            if (success)
            {
                resultDO.Response_code = "00";
                resultDO.Status = "3";

                return DskSuccessfulPayment(gid, vposRedirectGid, resultDO);
            }
            else
            {
                return DskFailedPayment(gid, vposRedirectGid, resultDO);
            }
        }

        [NonAction]
        private ActionResult AutoPostEservicePayment(string clientId, string secretKey, string postUrl, string paymentRequestIdentifier, Guid? vposResultGid, EserviceVposResultStatus status, string errorMessage)
        {
            EserviceVposResultDO eserviceVposResultDO = new EserviceVposResultDO();
            eserviceVposResultDO.RequestId = paymentRequestIdentifier;
            eserviceVposResultDO.VposResultGid = vposResultGid.HasValue ? vposResultGid.ToString() : null;
            eserviceVposResultDO.Status = status.ToString();
            eserviceVposResultDO.ErrorMessage = errorMessage;
            eserviceVposResultDO.ResultTime = Formatter.DateTimeToIso8601Format(DateTime.Now);

            var hmacAndCryptData = HmacRequestHelper.GetHmacAndCryptData(eserviceVposResultDO, secretKey, false);

            AutoPostVM model = new AutoPostVM();
            model.PostUrl = postUrl;
            model.PostValues = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<string, string>>();
            model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("clientId", clientId));
            model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("hmac", hmacAndCryptData.Item1));
            model.PostValues.Add(new System.Collections.Generic.KeyValuePair<string, string>("data", hmacAndCryptData.Item2));

            return View(MVC.Shared.Views._AutoPost, model);
        }

        [NonAction]
        private ActionResult AutoPostDsk(VposRequestDataVO vposRequestDataVO, Guid? vposRedirectGid)
        {
            if (vposRequestDataVO == null)
            {
                throw new Exception("Vpos request is invalid. Unable to get VposRequestDataVO data.");
            }

            AutoPostVM model = DskVposHelper.CreateAutoPostVM(vposRequestDataVO, vposRedirectGid);

            if (AppSettings.EPaymentsWeb_SimulateVposPayment)
            {
                Tuple<Guid, Guid?, string> resultData = new Tuple<Guid, Guid?, string>(
                    vposRequestDataVO.ЕserviceClientGid,
                    vposRedirectGid,
                    String.Format("{0}_{1}", vposRequestDataVO.PaymentRequestIdentifier, Guid.NewGuid().ToString().Substring(0, 8)));

                return View(MVC.Shared.Views._VposPaymentSumulation, resultData);
            }
            else
            {
                return View(MVC.Shared.Views._AutoPost, model);
            }
        }

        //Borica methods

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult BoricaPayment_(Guid id)
        {
            VposRequestDataVO dataVO;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                dataVO = this.webRepository.GetVposRequestDataByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                dataVO = this.webRepository.GetVposRequestDataByAccessCode(this.CurrentUser.AccessCode);
            }

            return AutoRedirectBorica(dataVO, null);
        }

        //Borica methods

        //[NonAction]
        //[Obsolete]
        //private ActionResult AutoRedirectBorica(VposRequestDataVO vposRequestDataVO, int? vposRedirectId)
        //{
        //    if (vposRequestDataVO == null)
        //    {
        //        throw new Exception("Vpos request is invalid. Unable to get VposRequestDataVO data.");
        //    }

        //    string boricaRequestIdentifier = this.webRepository.GenerateVposBoricaRequestIdentifier();

        //    VposBoricaRequest boricaRequest = new VposBoricaRequest();
        //    boricaRequest.RequestIdentifier = boricaRequestIdentifier;
        //    boricaRequest.PaymentRequestId = vposRequestDataVO.PaymentRequestId;
        //    boricaRequest.VposRedirectId = vposRedirectId;
        //    boricaRequest.RedirectUrl = BoricaVposHelper.CreateBOReqRedirectUrl(vposRequestDataVO, boricaRequestIdentifier);
        //    boricaRequest.CreateDate = DateTime.Now;

        //    this.systemRepository.AddEntity<VposBoricaRequest>(boricaRequest);
        //    this.unitOfWork.Save();

        //    return Redirect(boricaRequest.RedirectUrl);
        //}

        [NonAction]
        private ActionResult AutoRedirectBorica(VposRequestDataVO vposRequestDataVO, int? vposRedirectId)
        {
            if (vposRequestDataVO == null)
            {
                throw new Exception("Vpos request is invalid. Unable to get VposRequestDataVO data.");
            }

            EserviceClient eserviceClient =
                    this.unitOfWork.DbContext.Set<EserviceClient>()
                    .Where(e => e.Gid == vposRequestDataVO.ЕserviceClientGid)
                    .SingleOrDefault();

            if (vposRequestDataVO == null)
            {
                throw new Exception("EserviceClient not found.");
            }

            int boricaRequestIdentifier;

            using (var transaction = this.unitOfWork.BeginTransaction())
            {
                boricaRequestIdentifier = this.systemRepository.GetPaymentRequestTotal();
                transaction.Commit();
            }

            using (X509Certificate2 boricaCertificate = ReadBoricaCertificate(eserviceClient))
            {
                VposBoricaRequest boricaRequest = new VposBoricaRequest();
                boricaRequest.RequestIdentifier = boricaRequestIdentifier.ToString("000000");
                boricaRequest.PaymentRequestId = vposRequestDataVO.PaymentRequestId;
                boricaRequest.VposRedirectId = vposRedirectId;
                boricaRequest.CreateDate = DateTime.Now;

                APGWPaymentRequestDataVO apgwRequest = BoricaVposHelper.CreateBoricaRequestModel(
                    vposRequestDataVO,
                    boricaRequestIdentifier,
                    eserviceClient,
                    boricaCertificate);

                apgwRequest.P_Sign = this.CreatePSign(apgwRequest.GetPSignData(), boricaCertificate);

                boricaRequest.RedirectUrl = apgwRequest.PostUrl;

                this.systemRepository.AddEntity<VposBoricaRequest>(boricaRequest);
                this.unitOfWork.Save();

                return this.View(MVC.Shared.Views._AutoPostBoricaVpos, apgwRequest);
            }
        }

        [HttpGet]
        [Obsolete]
        [NonAction]
        public virtual ActionResult BoricaPayment(string eBorica)
        {
            //return Redirect("https://localhost:44300" + this.Request.RawUrl);

            BoricaVposResultDO resultDO = BoricaVposHelper.GetBoricaVposResultDOFromBOResp(eBorica);

            VposBoricaRequest boricaRequest = this.webRepository.GetVposBoricaRequestByRequestIdentifier(resultDO.RequestIdentifier);

            PaymentRequest paymentRequest = this.webRepository.GetPaymentRequestById(boricaRequest.PaymentRequestId);

            EserviceClient eserviceClient = this.systemRepository.GetEserviceClientById(paymentRequest.EserviceClientId);

            GlobalValue boricaVposResponseSignCertFileName = systemRepository.GetGlobalValueByKey(GlobalValueKey.BoricaVposResponseSignCertFileName);

            bool isRequestValid = BoricaVposHelper.IsBORespRequestValid(eBorica, boricaVposResponseSignCertFileName.Value);

            if (!isRequestValid)
            {
                throw new Exception("Vpos BoricaBOResp request is invalid.");
            }

            VposRedirect vposRedirect = null;
            if (boricaRequest.VposRedirectId.HasValue)
            {
                vposRedirect = this.systemRepository.GetVposRedirectById(boricaRequest.VposRedirectId.Value);
            }

            EserviceVposResultStatus resultStatus = resultDO.ResultCode == "00" ? EserviceVposResultStatus.Success : (resultDO.ResultCode == "94" ? EserviceVposResultStatus.CanceledByUser : EserviceVposResultStatus.Failure);

            return ProcessVposResultRequest(resultStatus, paymentRequest, eserviceClient, new JavaScriptSerializer().Serialize(resultDO), resultDO.RequestIdentifier, vposRedirect);
        }

        [HttpPost]
        public virtual ActionResult BoricaPayment(APGWPaymentResponseDataDO apgwResponse)
        {
            GlobalValue boricaVposResponseSignCertFileName = systemRepository.GetGlobalValueByKey(GlobalValueKey.BoricaVposResponseSignCertFileName);

            bool isRequestValid = BoricaVposHelper.IsBORespRequestValid(apgwResponse, boricaVposResponseSignCertFileName.Value);

            if (!isRequestValid)
            {
                throw new Exception("Vpos BoricaBOResp request is invalid.");
            }

            BoricaVposResultDO resultDO = BoricaVposHelper.GetBoricaVposResultDOFromBOResp(apgwResponse);

            DateTime transactionDate = DateTime.ParseExact(resultDO.DateTime, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

            DateTime startDate = transactionDate.AddHours(-1);
            DateTime endDate = startDate.AddDays(1);

            VposBoricaRequest boricaRequest = this.webRepository
                .GetVposBoricaRequestByRequestIdentifier(resultDO.RequestIdentifier, startDate, endDate);

            PaymentRequest paymentRequest = this.webRepository.GetPaymentRequestById(boricaRequest.PaymentRequestId);

            EserviceClient eserviceClient = this.systemRepository.GetEserviceClientById(paymentRequest.EserviceClientId);

            VposRedirect vposRedirect = null;
            if (boricaRequest.VposRedirectId.HasValue)
            {
                vposRedirect = this.systemRepository.GetVposRedirectById(boricaRequest.VposRedirectId.Value);
            }

            EserviceVposResultStatus resultStatus = resultDO.ResultCode == "00" ? EserviceVposResultStatus.Success : (resultDO.ResultCode == "-25" ? EserviceVposResultStatus.CanceledByUser : EserviceVposResultStatus.Failure);

            return ProcessVposResultRequest(resultStatus, paymentRequest, eserviceClient, new JavaScriptSerializer().Serialize(resultDO), resultDO.RequestIdentifier, vposRedirect);
        }

        //FiBank methods

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult FiBankPayment(Guid id)
        {
            VposRequestDataVO dataVO;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                dataVO = this.webRepository.GetVposRequestDataByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                dataVO = this.webRepository.GetVposRequestDataByAccessCode(this.CurrentUser.AccessCode);
            }

            return AutoPostFiBank(dataVO, null, null, this.Request.UserHostAddress);
        }

        [HttpPost]
        public virtual ActionResult FiBankPostResult(string trans_id, string ucaf_Cardholder_Confirm)
        {
            var vposBankRequest = this.unitOfWork.DbContext.Set<VposFiBankRequest>()
                .Where(e => e.TransactionId == trans_id)
                .SingleOrDefault();

            if (vposBankRequest == null)
            {
                throw new Exception("Fibank transaction id is not found");
            }

            var paymentRequest = this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Include(e => e.EserviceClient)
                .Include(p => p.ObligationType)
                .Where(e => e.PaymentRequestId == vposBankRequest.PaymentRequestId)
                .Single();

            vposBankRequest.IsVposPostCallbackReceived = true;
            vposBankRequest.VposPostCallbackReceiveDate = DateTime.Now;

            var transactionResult = FiBankVposHelper.CheckTransactionResult(
                paymentRequest.EserviceClient.FiBankVposAccessKeystoreFilename,
                paymentRequest.EserviceClient.FiBankVposAccessKeystorePassword,
                vposBankRequest.TransactionId,
                vposBankRequest.ClientIpAddress);

            vposBankRequest.IsPaymentSuccessful =
                transactionResult.Item1 == FiBankVposTransactionResult.Successful ? true :
                transactionResult.Item1 == FiBankVposTransactionResult.Unsuccessful ? false :
                (bool?)null;
            vposBankRequest.TransactionResult = transactionResult.Item2;
            vposBankRequest.TransactionResultReceiveDate = DateTime.Now;
            vposBankRequest.ResultStatus = VposRequestResultStatus.ReceivedByReturnUrlPost;

            this.unitOfWork.Save();

            if (transactionResult.Item1 == FiBankVposTransactionResult.Pending)
            {
                throw new Exception("Unabled to check fibank transaction result");
            }

            VposRedirect vposRedirect = null;
            if (vposBankRequest.VposRedirectId.HasValue)
            {
                vposRedirect = this.systemRepository.GetVposRedirectById(vposBankRequest.VposRedirectId.Value);
            }

            EserviceVposResultStatus resultStatus = vposBankRequest.IsPaymentSuccessful.Value ? EserviceVposResultStatus.Success : EserviceVposResultStatus.Failure;

            return ProcessVposResultRequest(resultStatus, paymentRequest, paymentRequest.EserviceClient, vposBankRequest.TransactionResult, vposBankRequest.TransactionId, vposRedirect);
        }

        [NonAction]
        private ActionResult AutoPostFiBank(VposRequestDataVO vposRequestDataVO, Guid? vposRedirectGid, int? vposRedirectId, string clientIpAddress)
        {
            if (vposRequestDataVO == null)
            {
                throw new Exception("Vpos request is invalid. Unable to get VposRequestDataVO data.");
            }

            string transactionId = FiBankVposHelper.GenerateTransactionId(vposRequestDataVO, clientIpAddress);

            if (String.IsNullOrWhiteSpace(transactionId))
            {
                throw new Exception("Failed generating Fibank transaction id.");
            }

            VposFiBankRequest vposBankRequest = new VposFiBankRequest();
            vposBankRequest.PaymentRequestId = vposRequestDataVO.PaymentRequestId;
            vposBankRequest.VposRedirectId = vposRedirectId;
            vposBankRequest.ClientIpAddress = clientIpAddress;
            vposBankRequest.TransactionId = transactionId;
            vposBankRequest.IsVposPostCallbackReceived = false;
            vposBankRequest.ResultStatus = VposRequestResultStatus.Pending;
            vposBankRequest.JobCheckResultFailedAttempts = 0;
            vposBankRequest.CreateDate = DateTime.Now;

            this.systemRepository.AddEntity<VposFiBankRequest>(vposBankRequest);
            this.unitOfWork.Save();

            AutoPostVM model = FiBankVposHelper.CreateAutoPostVM(vposBankRequest.TransactionId);

            if (AppSettings.EPaymentsWeb_SimulateVposPayment)
            {
                Tuple<Guid, Guid?, string> resultData = new Tuple<Guid, Guid?, string>(
                    vposRequestDataVO.ЕserviceClientGid,
                    vposRedirectGid,
                    String.Format("{0}_{1}", vposRequestDataVO.PaymentRequestIdentifier, Guid.NewGuid().ToString().Substring(0, 8)));

                return View(MVC.Shared.Views._VposPaymentSumulation, resultData);
            }
            else
            {
                return View(MVC.Shared.Views._AutoPost, model);
            }
        }

        //Epay methods

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult EpayPayment(Guid id)
        {
            VposRequestDataVO dataVO;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                dataVO = this.webRepository.GetVposRequestDataByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                dataVO = this.webRepository.GetVposRequestDataByAccessCode(this.CurrentUser.AccessCode);
            }

            return AutoPostEpay(dataVO, null, null, this.Request.UserHostAddress);
        }

        [HttpGet]
        public virtual ActionResult EpayCheckResult(string res, string ino)
        {
            VposEpayRequest vposRequest = this.webRepository.GetVposEpayRequestByInvoiceNo(ino.Trim());

            if (vposRequest == null)
            {
                throw new NotImplementedException("Invalid invoice number: " + ino);
            }

            bool? callbackSuccessConfirmation =
                res == null ? (bool?)null : (res.Trim().ToLower() == "true" ? true : ((res.Trim().ToLower() == "false" ? false : (bool?)null)));

            vposRequest.IsVposPostCallbackReceived = true;
            vposRequest.VposPostCallbackReceiveDate = DateTime.Now;
            vposRequest.VposPostCallbackSuccessConfirmation = callbackSuccessConfirmation;

            this.unitOfWork.Save();

            //begin checking for transaction result

            //VposEpayRequest vposRequestCheck = null;

            if (vposRequest.TransactionResult == null && callbackSuccessConfirmation != false)
            {
                for (int i = 0; i < 20; i++)
                {
                    vposRequest = this.webRepository.GetVposEpayRequestById(vposRequest.VposEpayRequestId);

                    if (vposRequest.TransactionResult != null)
                        break;

                    Thread.Sleep(2000);
                }
            }

            //end of checking

            //set resultStatus

            EserviceVposResultStatus resultStatus;

            if (callbackSuccessConfirmation == false)
            {
                resultStatus = EserviceVposResultStatus.CanceledByUser;
            }
            else if (vposRequest.IsPaymentSuccessful == true)
            {
                resultStatus = EserviceVposResultStatus.Success;
            }
            else
            {
                resultStatus = EserviceVposResultStatus.Failure;
            }

            vposRequest.VposPostCallbackAisRedirectedStatus = resultStatus.ToString();
            this.unitOfWork.Save();

            //process aid vpos request or return to EPayments list form

            EserviceClient eserviceClient = this.systemRepository.GetEserviceClientById(vposRequest.PaymentRequest.EserviceClientId);

            VposRedirect vposRedirect = null;
            if (vposRequest.VposRedirectId.HasValue)
            {
                vposRedirect = this.systemRepository.GetVposRedirectById(vposRequest.VposRedirectId.Value);
            }

            if (vposRedirect == null)
            {
                TempData[TempDataKeys.IsVposPaymentSuccessful] = resultStatus == EserviceVposResultStatus.Success;
                TempData[TempDataKeys.VposPaymentMessage] = resultStatus == EserviceVposResultStatus.Success ?
                    String.Format("Платихте успешно с карта задължение № {0}", vposRequest.PaymentRequest.PaymentRequestIdentifier) :
                    (resultStatus == EserviceVposResultStatus.CanceledByUser ? "Отказахте плащане с карта." : "Неуспешно плащане с карта.");

                return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }
            else
            {
                return AutoPostEservicePayment(
                    eserviceClient.ClientId,
                    eserviceClient.SecretKey,
                    resultStatus == EserviceVposResultStatus.CanceledByUser ? vposRedirect.CancelUrl : vposRedirect.OkUrl,
                    vposRequest.PaymentRequest.PaymentRequestIdentifier,
                    vposRedirect.Gid,
                    resultStatus,
                    null);//TODO: add errorMessage
            }
        }

        [HttpPost]
        public string EpayPaymentResult(string encoded, string checksum)
        {
            if (checksum == EpayVposHelper.CalculateHmac(AppSettings.EPaymentsWeb_EpayVposSecret, encoded))
            {
                string vposResultData = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));

                string[] arrVposResultData = vposResultData.Split(new string[] { "INVOICE" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < arrVposResultData.Length; i++)
                {
                    arrVposResultData[i] = "INVOICE" + arrVposResultData[i];
                }

                string response = "";
                foreach (var item in arrVposResultData)
                {
                    EpayResultData resultData = EpayVposHelper.GetEpayResultData(item);
                    string singleResponse = "";
                    try
                    {
                        singleResponse = HandleSingleEpayPayment(item, resultData);
                    }
                    catch (Exception ex)
                    {
                        this.logger.Error(ex);
                        singleResponse = "Error for data: " + item;
                    }

                    response += singleResponse;
                }

                return response;
            }
            else
            {
                return "invalid checksum";
            }
        }

        [NonAction]
        private string HandleSingleEpayPayment(string vposResultData, EpayResultData resultData)
        {
            VposEpayRequest vposRequest = this.webRepository.GetVposEpayRequestByInvoiceNo(resultData.Invoice);

            if (vposRequest == null)
            {
                return $"INVOICE={resultData.Invoice}:STATUS=invoice not found\n";
            }
            else
            {
                vposRequest.TransactionResult = vposResultData;
                vposRequest.TransactionResultReceiveDate = DateTime.Now;
                vposRequest.IsPaymentSuccessful = resultData.Status == EpayPaymentStatus.Paid;

                if (vposRequest.PaymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Pending)
                {
                    EserviceVposResultStatus resultStatus = resultData.Status.ToEserviceVposResultStatus();
                    PaymentRequest paymentRequest = vposRequest.PaymentRequest;
                    EserviceClient eserviceClient = this.systemRepository.GetEserviceClientById(paymentRequest.EserviceClientId);
                    string vposResultPostData = vposResultData;
                    string vposAuthorizationId = resultData.Invoice;

                    ProcessVposResult(
                        resultStatus,
                        paymentRequest,
                        eserviceClient,
                        vposResultPostData,
                        vposAuthorizationId,
                        false,
                        false);
                }
                else
                {
                    this.unitOfWork.Save();
                }

                return $"INVOICE={resultData.Invoice}:STATUS=OK\n";
            }
        }

        [NonAction]
        private ActionResult AutoPostEpay(VposRequestDataVO vposRequestDataVO, Guid? vposRedirectGid, int? vposRedirectId, string clientIpAddress)
        {
            if (vposRequestDataVO == null)
            {
                throw new Exception("Vpos request is invalid. Unable to get VposRequestDataVO data.");
            }

            PaymentRequest paymentRequest = this.webRepository.GetPaymentRequestById(vposRequestDataVO.PaymentRequestId);

            VposEpayRequest vposBankRequest = new VposEpayRequest();
            vposBankRequest.PaymentRequestId = vposRequestDataVO.PaymentRequestId;
            vposBankRequest.VposRedirectId = vposRedirectId;
            vposBankRequest.ClientIpAddress = clientIpAddress;
            vposBankRequest.TransactionInvoiceNo = GenerateInvoiceNo(vposRequestDataVO.PaymentRequestIdentifier);
            vposBankRequest.IsVposPostCallbackReceived = false;
            vposBankRequest.CreateDate = DateTime.Now;

            this.systemRepository.AddEntity<VposEpayRequest>(vposBankRequest);
            this.unitOfWork.Save();

            AutoPostVM model = EpayVposHelper.CreateAutoPostVM(vposBankRequest.TransactionInvoiceNo, paymentRequest);

            if (AppSettings.EPaymentsWeb_SimulateVposPayment)
            {
                Tuple<Guid, Guid?, string> resultData = new Tuple<Guid, Guid?, string>(
                    vposRequestDataVO.ЕserviceClientGid,
                    vposRedirectGid,
                    String.Format("{0}_{1}", vposRequestDataVO.PaymentRequestIdentifier, Guid.NewGuid().ToString().Substring(0, 8)));

                return View(MVC.Shared.Views._VposPaymentSumulation, resultData);
            }
            else
            {
                return View(MVC.Shared.Views._AutoPost, model);
            }
        }

        private string GenerateInvoiceNo(string paymentRequestIdentifier)
        {
            string invoiceNo = null;

            while (true)
            {
                invoiceNo = $"{GenerateEpayInvoicePrefix()}{paymentRequestIdentifier}";

                if (!this.webRepository.VposEpayInvoiceNumberExists(invoiceNo))
                    break;
            }

            return invoiceNo;
        }

        private static Random random = new Random();
        public static string GenerateEpayInvoicePrefix()
        {
            const string leadingSymbols = "123456789";
            const string symbols = "0123456789";

            return $"{new string(Enumerable.Repeat(leadingSymbols, 1).Select(s => s[random.Next(s.Length)]).ToArray())}{new string(Enumerable.Repeat(symbols, 5).Select(s => s[random.Next(s.Length)]).ToArray())}";
        }

        //DskEcomm methods

        [HttpGet]
        public virtual ActionResult DskEcommPaymentResult(string orderId)
        {
            var vposBankRequest = this.unitOfWork.DbContext.Set<VposDskEcommRequest>()
                .Where(e => e.TransactionId == orderId)
                .SingleOrDefault();

            if (vposBankRequest == null)
            {
                throw new Exception("DskEcomm transaction id is not found");
            }

            var paymentRequest = this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Include(e => e.EserviceClient.VposClient)
                .Include(p => p.ObligationType)
                .Where(e => e.PaymentRequestId == vposBankRequest.PaymentRequestId)
                .Single();

            vposBankRequest.IsVposPostCallbackReceived = true;
            vposBankRequest.VposPostCallbackReceiveDate = DateTime.Now;

            var transactionResult = DskEcommVposService.CheckOrderStatus(
                paymentRequest.EserviceClient.VposClient.PaymentRequestUrl,
                paymentRequest.EserviceClient.DskVposMerchantId,
                paymentRequest.EserviceClient.DskVposMerchantPassword,
                vposBankRequest.TransactionId);

            vposBankRequest.IsPaymentSuccessful =
                transactionResult.Item1.OrderStatus.HasValue && transactionResult.Item1.OrderStatus == DskEcommVposTransactionResult.Successful ? true :
                transactionResult.Item1.OrderStatus.HasValue && transactionResult.Item1.OrderStatus != DskEcommVposTransactionResult.Successful ? false :
                (bool?)null;

            vposBankRequest.TransactionResult = transactionResult.Item2;
            vposBankRequest.TransactionResultReceiveDate = DateTime.Now;
            vposBankRequest.ResultStatus = VposRequestResultStatus.ReceivedByReturnUrlPost;

            this.unitOfWork.Save();

            if (transactionResult.Item1.OrderStatus == DskEcommVposTransactionResult.Pending)
            {
                throw new Exception("Unabled to check DskEcomm transaction result");
            }

            VposRedirect vposRedirect = null;
            if (vposBankRequest.VposRedirectId.HasValue)
            {
                vposRedirect = this.systemRepository.GetVposRedirectById(vposBankRequest.VposRedirectId.Value);
            }

            EserviceVposResultStatus resultStatus = vposBankRequest.IsPaymentSuccessful.Value ? EserviceVposResultStatus.Success : EserviceVposResultStatus.Failure;

            return ProcessVposResultRequest(resultStatus, paymentRequest, paymentRequest.EserviceClient, vposBankRequest.TransactionResult, vposBankRequest.TransactionId, vposRedirect);
        }


        [NonAction]
        private ActionResult AutoPostDskEcomm(VposRequestDataVO vposRequestDataVO, Guid? vposRedirectGid, int? vposRedirectId, string clientIpAddress)
        {
            if (vposRequestDataVO == null)
            {
                throw new Exception("Vpos request is invalid. Unable to get VposRequestDataVO data.");
            }

            DskEcommVposRegisteredOrder registeredOrder = DskEcommVposService.RegisterOrder(
                vposRequestDataVO.VposPaymentRequestUri,
                vposRequestDataVO.DskVposMerchantId,
                vposRequestDataVO.DskVposMerchantPassword,
                vposRequestDataVO.PaymentRequestIdentifier,
                vposRequestDataVO.PaymentAmount,
                vposRequestDataVO.PaymentReason);

            if (String.IsNullOrWhiteSpace(registeredOrder.OrderId) || 
                (registeredOrder.ErrorCode != null && !registeredOrder.ErrorCode.Equals("0")) || 
                registeredOrder.ErrorMessage != null)
            {
                throw new Exception($"Failed generating DskEcomm transaction id. ErrorCode: {registeredOrder.ErrorCode ?? string.Empty}; ErrorMessage: {registeredOrder.ErrorMessage ?? string.Empty}");
            }

            VposDskEcommRequest vposBankRequest = new VposDskEcommRequest();
            vposBankRequest.PaymentRequestId = vposRequestDataVO.PaymentRequestId;
            vposBankRequest.VposRedirectId = vposRedirectId;
            vposBankRequest.ClientIpAddress = clientIpAddress;
            vposBankRequest.TransactionId = registeredOrder.OrderId;
            vposBankRequest.IsVposPostCallbackReceived = false;
            vposBankRequest.ResultStatus = VposRequestResultStatus.Pending;
            vposBankRequest.JobCheckResultFailedAttempts = 0;
            vposBankRequest.CreateDate = DateTime.Now;

            this.systemRepository.AddEntity<VposDskEcommRequest>(vposBankRequest);
            this.unitOfWork.Save();

            if (AppSettings.EPaymentsWeb_SimulateVposPayment)
            {
                Tuple<Guid, Guid?, string> resultData = new Tuple<Guid, Guid?, string>(
                    vposRequestDataVO.ЕserviceClientGid,
                    vposRedirectGid,
                    String.Format("{0}_{1}", vposRequestDataVO.PaymentRequestIdentifier, Guid.NewGuid().ToString().Substring(0, 8)));

                return View(MVC.Shared.Views._VposPaymentSumulation, resultData);
            }
            else
            {
                return Redirect(registeredOrder.FormUrl);
            }
        }

        [NonAction]
        public ActionResult BoricaPayments(Guid[] guids) 
        {
            if (guids == null || guids.Length == 0)
            {
                return this.RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }

            List<PaymentRequest> paymentRequests = this.webRepository.GetPendingPaymentRequestByUid(guids);
            if (paymentRequests.Count != guids.Length)
            {
                TempData[TempDataKeys.IsVposPaymentSuccessful] = false;
                TempData[TempDataKeys.VposPaymentMessage] = string.Format("Някои от вашите задължения не са в статус \"{0}\".",
                    PaymentRequestStatus.Pending.GetDescription());

                return this.RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }

            DateTime nowDate = DateTime.Now;
            int counter;

            foreach (var pr in paymentRequests)
            {
                pr.PaymentRequestStatusId = PaymentRequestStatus.InProcess;
                pr.PaymentRequestStatusChangeTime = DateTime.Now;

                foreach (var transaction in pr.BoricaTransactions) 
                {
                    if (transaction.Rc == "00" && transaction.Action == "0")
                    {
                        pr.PaymentRequestStatusId = PaymentRequestStatus.Paid;
                        pr.PaymentRequestStatusChangeTime = DateTime.Now;
                        this.unitOfWork.Save();

                        TempData[TempDataKeys.IsVposPaymentSuccessful] = false;
                        TempData[TempDataKeys.VposPaymentMessage] = string.Format("Някои от вашите избрани задължения вече са били платени.");

                        return this.RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
                    }
                }

                this.unitOfWork.Save();
            }

            using (var transaction = this.unitOfWork.BeginTransaction())
            {
                counter = this.systemRepository.GetPaymentRequestTotal();
                transaction.Commit();
            }

            BoricaTransaction boricaTransaction = new BoricaTransaction()
            {
                Amount = paymentRequests.Sum(pr => pr.PaymentAmount),
                Order = String.Format("{0}{1}", Formatter.DateToShortFormat(nowDate), counter),
                PaymentRequests = paymentRequests,
                TransactionDate = nowDate,
                Gid = Guid.NewGuid(),
                TransactionStatusId = (int)BoricaTransactionStatusEnum.Pending
            };

            this.unitOfWork.DbContext.Set<BoricaTransaction>().Add(boricaTransaction);

            this.unitOfWork.Save();

            APGWPaymentRequestDataVO apgwRequest = new APGWPaymentRequestDataVO()
            {
                Merchant_Name = AppSettings.EPaymentsWeb_MerchantName,
                Merchant = AppSettings.EPaymentsWeb_CentralVposMerchantId,
                DESC = AppSettings.EPaymentsWeb_CentralVposDescription,
                Amount = boricaTransaction.Amount,
                Email = AppSettings.EPaymentsWeb_MerchantEmail,
                Merchant_Url = AppSettings.EPaymentsWeb_MerchantUrl,
                Order = counter.ToString("000000"),
                AdCustomBorOrderId = counter.ToString("000000") + AppSettings.EPaymentsWeb_CentralVposPrefixHelper + counter.ToString("00000000000000"),
                ADDENDUM = AppSettings.EPaymentsWeb_CentralVposADDENDUM,
                Date = boricaTransaction.TransactionDate.ToUniversalTime(),
                Terminal = AppSettings.EPaymentsWeb_CentralVposDevTerminalId
            };

            string pSignData = apgwRequest.GetPSignData();
            this.logger.Info(apgwRequest.ToString() + ", pSignData: " + pSignData);
            apgwRequest.P_Sign = this.CreatePSign(pSignData);

            return this.View(MVC.Shared.Views._AutoPostBoricaVpos, apgwRequest);
        }

        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        [HttpGet]
        public virtual ActionResult Borica(Guid[] gids)
        {
            return this.BoricaPayments(gids);
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult PayWithBorica(AuthRequestDO requestDO)
        {
            string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
            JObject dataJObject = JObject.Parse(dataJson);

            string paymentRequestIdentifier = dataJObject.GetValue("paymentRequestIdentifier", StringComparison.OrdinalIgnoreCase).Value<string>();
            string redirectUrl = dataJObject.GetValue("redirectUrl", StringComparison.OrdinalIgnoreCase).Value<string>();
            var paymentRequest = this.webRepository.GetPaymentRequestByIdentifier(paymentRequestIdentifier);

            if (paymentRequest == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest, "payment request not found");
            }

            if (paymentRequest.PaymentRequestStatusId != PaymentRequestStatus.Pending)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest, $"payment request must be in Pending status. Current Status: {paymentRequest.PaymentRequestStatusId}");
            }

            if (!string.IsNullOrEmpty(redirectUrl))
            {
                paymentRequest.RedirectUrl = redirectUrl;
                this.unitOfWork.Save();
            }
            return BoricaPayments(new Guid[] { paymentRequest.Gid });
        }

        [HttpPost]
        public virtual ActionResult BoricaBackref(APGWPaymentResponseDataDO apgwResponse)
        {
            string MasterCardSoftRc = "65";
            string VisaCardSoftRc = "1A";
            string redirectUrl = string.Empty;

            if (apgwResponse.IsResponseValid())
            {
                logger.Info("apgwResponse is valid.");
                byte[] signature = BinHexHelper.HexStringToByteArray(apgwResponse.P_Sign);

                var data = Encoding.UTF8.GetBytes(apgwResponse.GetPSignData());
                var certificate = BoricaCvposHelper.GetBoricaCvposDev();

                using (var rsa = certificate.GetRSAPublicKey())
                {
                    var isValid = rsa.VerifyData(data, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                    if (isValid)
                    {
                        logger.Info("apgwResponse data is verifyed with certificate's RSA.");
                        var transactionOrder = String.Format("{0}{1}", Formatter.DateToShortFormat(apgwResponse.Date), apgwResponse.Order.TrimStart(new char[] { '0' }));
                 
                        
                        var boricaTransaction = this.unitOfWork.DbContext.Set<BoricaTransaction>()
                            .Include(bt => bt.PaymentRequests.Select(pr => pr.ObligationType))
                            .FirstOrDefault(bt => bt.Order == transactionOrder);

                        redirectUrl = boricaTransaction.PaymentRequests.FirstOrDefault().RedirectUrl;

                        if (apgwResponse.Action == 0)
                        {
                            if (boricaTransaction == null)
                            {
                                TempData[TempDataKeys.IsVposPaymentSuccessful] = false;
                                TempData[TempDataKeys.VposPaymentMessage] = "Невалидна транзакция.";

                                return this.RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
                            }

                            switch (apgwResponse.TrType)
                            {
                                case TrTypeEnum.Payment:
                                    if (string.Equals(apgwResponse.Rc, MasterCardSoftRc, StringComparison.OrdinalIgnoreCase) || string.Equals(apgwResponse.Rc, VisaCardSoftRc, StringComparison.OrdinalIgnoreCase))
                                    {
                                        boricaTransaction = apgwResponse.UpdateBoricaTransactionCanceledPayment(boricaTransaction);
                                        this.unitOfWork.Save();

                                        return this.SoftBoricaRequest(apgwResponse, boricaTransaction.PaymentRequests);
                                    }

                                    if (apgwResponse.Rc == "00")
                                    {
                                        boricaTransaction = apgwResponse.UpdateBoricaTransactionSuccessPayment(boricaTransaction);

                                        IEnumerable<string> identificators = boricaTransaction
                                            .PaymentRequests.Select(pr => pr.ApplicantUin);

                                        List<User> users = this.unitOfWork.DbContext.Set<User>()
                                            .Where(u => identificators.Contains(u.Egn))
                                            .ToList();

                                        foreach (var pr in boricaTransaction.PaymentRequests)
                                        {
                                            pr.ObligationStatusId = ObligationStatusEnum.IrrevocableOrder;
                                            pr.PaymentRequestStatusId = PaymentRequestStatus.Paid;
                                            pr.PaymentRequestStatusChangeTime = DateTime.Now;

                                            User user = users.FirstOrDefault(u => u.Egn == pr.ApplicantUin);

                                            if (user != null && !String.IsNullOrWhiteSpace(user.Email))
                                            {
                                                if (user.StatusObligationNotifications)
                                                {
                                                    Email email = user.CreateStatusObligationNotificationEmail(pr);

                                                    this.systemRepository.AddEntity<Email>(email);
                                                }

                                                if (user.StatusNotifications)
                                                {
                                                    Email email = user.CreateStatusNotificationEmail(pr);

                                                    this.systemRepository.AddEntity<Email>(email);
                                                }
                                            }

                                            if (!String.IsNullOrWhiteSpace(pr.AdministrativeServiceNotificationURL))
                                            {
                                                EserviceNotification statusNotification = new EserviceNotification(pr);

                                                this.systemRepository.AddEntity<EserviceNotification>(statusNotification);
                                            }
                                        }

                                        this.unitOfWork.Save();

                                        TempData[TempDataKeys.IsVposPaymentSuccessful] = true;
                                        TempData[TempDataKeys.VposPaymentMessage] = string.Format("Платихте успешно Вашите задължения с карта {0}", apgwResponse.Card);
                                        
                                        if (!string.IsNullOrEmpty(redirectUrl))
                                        {
                                            return Redirect(redirectUrl);
                                        }

                                        return this.RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
                                    }
                                    break;
                            }
                        }
                    }
                }
            }

            string defaultErrorMessage = "Грешка при плащане с карта.";
            
            switch (apgwResponse.TrType)
            {
                case TrTypeEnum.Payment:
                    if (int.TryParse(apgwResponse.Rc, out int responseCode))
                    {
                        if (responseCode >= 0)
                        {
                            var responseError = (CvposResponseCodesEnum)(responseCode);
                            var responseErrorDescription = responseError.GetDescription();
                            TempData[TempDataKeys.IsVposPaymentSuccessful] = false;
                            TempData[TempDataKeys.VposPaymentMessage] = string.Format("Грешка при плащане с карта. За повече информация, свържете се с вашата банка (Код {0}, Грешка {1} / {2})",
                                responseCode, apgwResponse.StatusMsg, responseErrorDescription);
                        }
                        else
                        {
                            TempData[TempDataKeys.IsVposPaymentSuccessful] = false;
                            TempData[TempDataKeys.VposPaymentMessage] = defaultErrorMessage;
                        }
                    }
                    else
                    {
                        TempData[TempDataKeys.IsVposPaymentSuccessful] = false;
                        TempData[TempDataKeys.VposPaymentMessage] = defaultErrorMessage;
                    }

                    return this.RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
                default:
                    TempData[TempDataKeys.IsVposPaymentSuccessful] = false;
                    TempData[TempDataKeys.VposPaymentMessage] = defaultErrorMessage;

                    return this.RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }
        }

        [NonAction]
        private ActionResult SoftBoricaRequest(APGWPaymentResponseDataDO apgwResponse, 
            ICollection<PaymentRequest> paymentRequests)
        {
            DateTime nowDate = DateTime.Now;
            int counter;

            paymentRequests = paymentRequests
                .Where(pr => pr.PaymentRequestStatusId == PaymentRequestStatus.Pending)
                .ToList();

            using (var transaction = this.unitOfWork.BeginTransaction())
            {
                counter = this.systemRepository.GetPaymentRequestTotal();
                transaction.Commit();
            }

            BoricaTransaction boricaTransaction = new BoricaTransaction()
            {
                Amount = paymentRequests.Sum(pr => pr.PaymentAmount),
                Order = String.Format("{0}{1}", Formatter.DateToShortFormat(nowDate), counter),
                PaymentRequests = paymentRequests,
                TransactionDate = DateTime.Now.ToUniversalTime(),
                Gid = Guid.NewGuid(),
                TransactionStatusId = (int)BoricaTransactionStatusEnum.Pending
            };

            this.unitOfWork.DbContext.Set<BoricaTransaction>().Add(boricaTransaction);

            this.unitOfWork.Save();

            APGWSoftRequestDataVO apgwRequest = new APGWSoftRequestDataVO()
            {
                Merchant_Name = AppSettings.EPaymentsWeb_MerchantName,
                Merchant = AppSettings.EPaymentsWeb_CentralVposMerchantId,
                DESC = AppSettings.EPaymentsWeb_CentralVposDescription,
                Amount = paymentRequests.Sum(pr => pr.PaymentAmount),
                Email = AppSettings.EPaymentsWeb_MerchantEmail,
                Merchant_Url = AppSettings.EPaymentsWeb_MerchantUrl,
                Order = counter.ToString("000000"),
                AdCustomBorOrderId = counter.ToString("000000") + AppSettings.EPaymentsWeb_CentralVposPrefixHelper + counter.ToString("00000000000000"),
                ADDENDUM = AppSettings.EPaymentsWeb_CentralVposADDENDUM,
                Date = boricaTransaction.TransactionDate,
                Terminal = AppSettings.EPaymentsWeb_CentralVposDevTerminalId
            };

            apgwRequest.P_Sign = this.CreatePSign(apgwRequest.GetPSignData());

            return this.View(MVC.Shared.Views._AutoPostBoricaVpos, apgwRequest);
        }

        [NonAction]
        private ActionResult SoftBoricaAuthorization(APGWPaymentResponseDataDO apgwResponse)
        {
            APGWSoftFirstAutorizationDataVO apgwRequest = new APGWSoftFirstAutorizationDataVO()
            {
                Merchant_Name = AppSettings.EPaymentsWeb_MerchantName,
                Merchant = AppSettings.EPaymentsWeb_CentralVposMerchantId,
                DESC = AppSettings.EPaymentsWeb_CentralVposDescription,
                Amount = apgwResponse.AmountValue ?? 0,
                Email = AppSettings.EPaymentsWeb_MerchantEmail,
                Merchant_Url = AppSettings.EPaymentsWeb_MerchantUrl,
                Order = apgwResponse.Order,
                AdCustomBorOrderId = apgwResponse?.AdCustomBorOrderId ?? apgwResponse.Order + AppSettings.EPaymentsWeb_CentralVposPrefixHelper + (int.Parse(apgwResponse.Order.Trim('0'))).ToString("00000000000000"),
                ADDENDUM = AppSettings.EPaymentsWeb_CentralVposADDENDUM,
                Date = apgwResponse.Date,
                Terminal = AppSettings.EPaymentsWeb_CentralVposDevTerminalId
            };

            apgwRequest.P_Sign = this.CreatePSign(apgwRequest.GetPSignData());

            return this.View(MVC.Shared.Views._AutoPostBoricaVpos, apgwRequest);
        }

        [NonAction]
        private ActionResult CreateEndingAuthorization(APGWPaymentResponseDataDO apgwResponse)
        {
            APGWEndingAuthorizationDataVO apgwRequest = new APGWEndingAuthorizationDataVO()
            {
                Amount = apgwResponse.AmountValue ?? 0,
                RRN = apgwResponse.Rrn,
                INT_REF = apgwResponse.INT_REF,
                DESC = AppSettings.EPaymentsWeb_CentralVposDescription,
                Terminal = AppSettings.EPaymentsWeb_CentralVposDevTerminalId,
                Merchant_Name = AppSettings.EPaymentsWeb_MerchantName,
                Merchant_Url = AppSettings.EPaymentsWeb_MerchantUrl,
                Merchant = AppSettings.EPaymentsWeb_CentralVposMerchantId,
                Email = AppSettings.EPaymentsWeb_MerchantEmail,
                Order = apgwResponse.Order,
                AdCustomBorOrderId = apgwResponse?.AdCustomBorOrderId ?? apgwResponse.Order + AppSettings.EPaymentsWeb_CentralVposPrefixHelper + (int.Parse(apgwResponse.Order.Trim('0'))).ToString("00000000000000"),
                ADDENDUM = AppSettings.EPaymentsWeb_CentralVposADDENDUM,
                Date = apgwResponse.Date,
            };

            apgwRequest.P_Sign = this.CreatePSign(apgwRequest.GetPSignData());

            return this.View(MVC.Shared.Views._AutoPostBoricaVpos, apgwRequest);
        }

        [NonAction]
        public virtual ActionResult BoricaAuthorization()
        {
            List<PaymentRequest> paymentRequests = this.webRepository.GetPendingPaymentRequestByUid(new Guid[] { new Guid("9BCE8AFF-3F33-4741-B2AD-D96E4B57F0DB") });

            DateTime nowDate = DateTime.Now;
            int counter;

            using (var transaction = this.unitOfWork.BeginTransaction())
            {
                counter = this.systemRepository.GetPaymentRequestCounter(nowDate);

                transaction.Commit();
            }

            BoricaTransaction boricaTransaction = new BoricaTransaction()
            {
                Amount = paymentRequests.Sum(pr => pr.PaymentAmount),
                Order = String.Format("{0}{1}", Formatter.DateToShortFormat(nowDate), counter),
                PaymentRequests = paymentRequests,
                TransactionDate = DateTime.Now.ToUniversalTime(),
                Gid = Guid.NewGuid(),
                TransactionStatusId = (int)BoricaTransactionStatusEnum.Pending
            };

            this.unitOfWork.DbContext.Set<BoricaTransaction>().Add(boricaTransaction);

            this.unitOfWork.Save();

            APGWFirstAuthorizationDataVO apgwRequest = new APGWFirstAuthorizationDataVO()
            {
                Amount = boricaTransaction.Amount,
                DESC = AppSettings.EPaymentsWeb_CentralVposDescription,
                Terminal = AppSettings.EPaymentsWeb_CentralVposDevTerminalId,
                Merchant_Name = AppSettings.EPaymentsWeb_MerchantName,
                Merchant_Url = AppSettings.EPaymentsWeb_MerchantUrl,
                Merchant = AppSettings.EPaymentsWeb_CentralVposMerchantId,
                Email = AppSettings.EPaymentsWeb_MerchantEmail,
                Order = counter.ToString("000000"),
                AdCustomBorOrderId = counter.ToString("000000") + AppSettings.EPaymentsWeb_CentralVposPrefixHelper + counter.ToString("00000000000000"),
                ADDENDUM = AppSettings.EPaymentsWeb_CentralVposADDENDUM,
                Date = boricaTransaction.TransactionDate,
            };

            apgwRequest.P_Sign = this.CreatePSign(apgwRequest.GetPSignData());

            return this.View(MVC.Shared.Views._AutoPostBoricaVpos, apgwRequest);
        }

        [NonAction]
        private ActionResult ProcessVposResultRequest(
            EserviceVposResultStatus resultStatus,
            PaymentRequest paymentRequest,
            EserviceClient eserviceClient,
            string vposResultPostData,
            string vposAuthorizationId,
            VposRedirect vposRedirect)
        {
            Guid? vposResultGid = ProcessVposResult(
                resultStatus,
                paymentRequest,
                eserviceClient,
                vposResultPostData,
                vposAuthorizationId,
                vposRedirect != null,
                true);

            if (vposRedirect == null)
            {
                TempData[TempDataKeys.IsVposPaymentSuccessful] = resultStatus == EserviceVposResultStatus.Success;
                TempData[TempDataKeys.VposPaymentMessage] = resultStatus == EserviceVposResultStatus.Success ?
                    String.Format("Платихте успешно с карта задължение № {0}", paymentRequest.PaymentRequestIdentifier) :
                    (resultStatus == EserviceVposResultStatus.CanceledByUser ? "Отказахте плащане с карта." : "Неуспешно плащане с карта.");

                return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
            }
            else
            {
                return AutoPostEservicePayment(
                    eserviceClient.ClientId,
                    eserviceClient.SecretKey,
                    resultStatus == EserviceVposResultStatus.CanceledByUser ? vposRedirect.CancelUrl : vposRedirect.OkUrl,
                    paymentRequest.PaymentRequestIdentifier,
                    vposResultGid.Value,
                    resultStatus,
                    null);//TODO: add errorMessage
            }
        }

        [NonAction]
        private Guid? ProcessVposResult(
            EserviceVposResultStatus resultStatus,
            PaymentRequest paymentRequest,
            EserviceClient eserviceClient,
            string vposResultPostData,
            string vposAuthorizationId,
            bool hasVposRedirect,
            bool createVposResult)
        {
            VposResult vposResult = null;

            if (createVposResult)
            {
                vposResult = new VposResult();
                vposResult.Gid = Guid.NewGuid();
                vposResult.PaymentRequestId = paymentRequest.PaymentRequestId;
                vposResult.PostData = vposResultPostData;
                vposResult.PostUrl = this.Request.RawUrl;
                vposResult.IsPaymentSuccessfull = resultStatus == EserviceVposResultStatus.Success;
                vposResult.IsPaymentCanceledByUser = resultStatus == EserviceVposResultStatus.CanceledByUser;
                vposResult.RequestDate = DateTime.Now;

                this.webRepository.AddEntity<VposResult>(vposResult);
            }

            if (resultStatus == EserviceVposResultStatus.Success)
            {
                paymentRequest.VposResult = vposResult;
                paymentRequest.IsVposAuthorized = true;
                paymentRequest.VposAuthorizationId = vposAuthorizationId;
                paymentRequest.PaymentRequestStatusId = Model.Enums.PaymentRequestStatus.Authorized;
                paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;
                paymentRequest.ObligationStatusId = Model.Enums.ObligationStatusEnum.IrrevocableOrder;
            }

            this.unitOfWork.Save();

            if (resultStatus == EserviceVposResultStatus.Success)
            {
                if (AppSettings.EPaymentsEventRegister_Enabled)
                {
                    EventRegisterNotification notification = new EventRegisterNotification(
                        paymentRequest.PaymentRequestId,
                        EventRegisterNotificationType.VposPaymentAuthorized,
                        String.Format("(еПлащане) Плащането на заявка с №{0} е авторизирано от виртуален ПОС", paymentRequest.PaymentRequestIdentifier),
                        paymentRequest.PaymentRequestIdentifier);

                    this.systemRepository.AddEntity<EventRegisterNotification>(notification);
                    this.unitOfWork.Save();
                }

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
                    if (hasVposRedirect)
                    {
                        EserviceNotification statusNotification = new EserviceNotification(paymentRequest);

                        this.systemRepository.AddEntity<EserviceNotification>(statusNotification);
                        this.unitOfWork.Save();
                    }
                    else
                    {
                        Task.Factory.StartNew(() =>
                        {
                            SendAuthorizedStatusNotificationAsync(paymentRequest, eserviceClient).RunAndForget();
                        });
                    }
                }
            }

            return createVposResult ? vposResult.Gid : (Guid?)null;
        }

        //StatusChange notifications

        [NonAction]
        private async Task SendAuthorizedStatusNotificationAsync(PaymentRequest request, EserviceClient еserviceClient)
        {
            EserviceNotification notification = new EserviceNotification(request, null);

            try
            {
                string postMediaType = "application/x-www-form-urlencoded";
                string acceptHeaderMediaType1 = "application/json";
                string acceptHeaderMediaType2 = "text/plain";

                var jsonDataBytes = Encoding.UTF8.GetBytes(notification.PostData);

                var base64Data = Convert.ToBase64String(jsonDataBytes);

                var hmac = HmacRequestHelper.CalculateHmac(еserviceClient.SecretKey, base64Data);

                Uri notificationUri = new Uri(notification.Url);

                string baseAddress = String.Format("{0}://{1}:{2}", notificationUri.Scheme, notificationUri.Host, notificationUri.Port);
                string postUri = notificationUri.PathAndQuery;

                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseAddress);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(acceptHeaderMediaType1));
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(acceptHeaderMediaType2));

                    var encodedHttpBody = String.Format("ClientId={0}&Hmac={1}&Data={2}", HttpUtility.UrlEncode(еserviceClient.ClientId), HttpUtility.UrlEncode(hmac), HttpUtility.UrlEncode(base64Data));

                    ServicePointManager.ServerCertificateValidationCallback += delegate (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                    {
                        return true;
                    };

                    try
                    {
                        var task = client.PostAsync(postUri, new StringContent(encodedHttpBody, Encoding.UTF8, postMediaType));

                        await task.ConfigureAwait(false);

                        string resultContent;
                        if (IsNotificationAccepted(task.Result, out resultContent))
                        {
                            notification.SetStatus(NotificationStatus.Sent);
                        }
                        else
                        {
                            string errorMessage = String.Format("Invalid response content. Status code - {0}, ResultContent - {1}", task.Result.StatusCode.ToString(), resultContent);
                            throw new Exception(errorMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        var exception = "HttpClientException: " + ex.Message ?? String.Empty + (ex.InnerException != null ? " Inner exception: " + ex.InnerException.Message ?? String.Empty : String.Empty);

                        notification.IncrementFailedAttempts(exception);
                        DateTime? sendNotBefore = CalculateNextSendingAttemptTime(notification.FailedAttempts);
                        notification.SetNextSendingAttemptTime(sendNotBefore, !sendNotBefore.HasValue);
                    }
                }
            }
            catch (Exception ex)
            {
                notification.SetStatus(NotificationStatus.Error);
                notification.IncrementFailedAttempts(ex.Message);
            }

            this.systemRepository.AddEntity<EserviceNotification>(notification);
            this.unitOfWork.Save();
        }

        [NonAction]
        private bool IsNotificationAccepted(HttpResponseMessage responseMessage, out string resultContent)
        {
            bool returnValue = false;
            resultContent = String.Empty;
            try
            {
                if (responseMessage.IsSuccessStatusCode)
                {
                    resultContent = responseMessage.Content.ReadAsStringAsync().Result;

                    JObject responseJson = JObject.Parse(resultContent);
                    returnValue = bool.Parse((string)responseJson.Properties().ToList().Single(e => e.Name.ToLower() == "success".ToLower()).Value);
                }

                return returnValue;
            }
            catch
            {
                if (resultContent.Length > 1000)
                {
                    resultContent = resultContent.Substring(0, 1000) + "... value is trimmed";
                }
                return returnValue;
            }
        }

        [NonAction]
        private DateTime? CalculateNextSendingAttemptTime(int failedAttempts)
        {
            DateTime? nextAttemptTime = null;

            if (failedAttempts < 6)
            {
                nextAttemptTime = DateTime.Now.AddMinutes(1);
            }
            else if (failedAttempts < 10)
            {
                nextAttemptTime = DateTime.Now.AddMinutes(15);
            }
            else if (failedAttempts < 15)
            {
                nextAttemptTime = DateTime.Now.AddHours(1);
            }
            else if (failedAttempts < 21)
            {
                nextAttemptTime = DateTime.Now.AddHours(3);
            }
            else if (failedAttempts < 25)
            {
                nextAttemptTime = DateTime.Now.AddHours(6);
            }

            return nextAttemptTime;
        }

        private string CreatePSign(string pSignData)
        {
            using (var certificate = new X509Certificate2(BoricaCvposHelper.GetBoricaCvposDevPfxPath(),
                AppSettings.EPaymentsWeb_CentralVposPrivateKeyPassphrase,
                X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable))
            {
                using (RSA rsa = certificate.GetRSAPrivateKey())
                {
                    var pSignBytes = Encoding.UTF8.GetBytes(pSignData);
                    byte[] signature = rsa.SignData(pSignBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                    var isValid = rsa.VerifyData(pSignBytes, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                    if (isValid == false)
                    {
                        throw new Exception("Грешка при валидация на подписаните на данни при изпращане към ЦВПОС.");
                    }

                    return BinHexHelper.ByteArrayToHexString(signature).ToUpper();
                }
            }
        }

        private X509Certificate2 ReadBoricaCertificate(EserviceClient eserviceClient)
        {
            if (eserviceClient == null)
            {
                throw new Exception("Не са подадени валидни параметри.");
            }

            string boricaCertificatePath = Path.Combine(AppSettings.EPaymentsWeb_BoricaCertificateFolder, eserviceClient.BoricaVposRequestSignCertFileName);

            if (!System.IO.File.Exists(boricaCertificatePath))
            {
                throw new Exception("Не е намерен сертификата.");
            }

            return new X509Certificate2(boricaCertificatePath,
                eserviceClient.BoricaVposRequestSignCertPassword,
                AppSettings.EPaymentsCommon_UseMachineKeySet == false ?
                X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable :
                X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable);
        }

        private string CreatePSign(string pSignData, X509Certificate2 certificate)
        {
            using (RSA rsa = certificate.GetRSAPrivateKey())
            {
                var pSignBytes = Encoding.UTF8.GetBytes(pSignData);
                byte[] signature = rsa.SignData(pSignBytes, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                var isValid = rsa.VerifyData(pSignBytes, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                if (isValid == false)
                {
                    throw new Exception("Грешка при валидация на подписаните на данни при изпращане към ЦВПОС.");
                }

                return BinHexHelper.ByteArrayToHexString(signature).ToUpper();
            }
        }
    }

    public static class TasksExtensions
    {
        public static void RunAndForget(this Task task) { }
    }
}