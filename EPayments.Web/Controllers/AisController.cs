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

namespace EPayments.Web.Controllers
{
    public partial class AisController : BaseController
    {
        private IUnitOfWork unitOfWork;
        private IWebRepository webRepository;
        private ISystemRepository systemRepository;

        public AisController(IUnitOfWork unitOfWork, IWebRepository webRepository, ISystemRepository systemRepository)
        {
            this.unitOfWork = unitOfWork;
            this.webRepository = webRepository;
            this.systemRepository = systemRepository;
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult PaymentOrder(AuthRequestDO requestDO)
        {
            AisRequestDO aisRequestDO = ValidateAndGetAisRequestDO(requestDO);

            var paymentOrderVO = this.webRepository.GetPaymentOrderByGidAndUin(aisRequestDO.PaymentRequest.Gid, aisRequestDO.PaymentRequest.ApplicantUin);

            OrderVM model = new OrderVM(paymentOrderVO, requestDO);

            return PartialView(MVC.Shared.Views._Order, model);
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual ActionResult PrintOrder(AuthRequestDO requestDO)
        {
            AisRequestDO aisRequestDO = ValidateAndGetAisRequestDO(requestDO);

            var paymentOrderVO = this.webRepository.GetPaymentOrderByGidAndUin(aisRequestDO.PaymentRequest.Gid, aisRequestDO.PaymentRequest.ApplicantUin);

            OrderVM model = new OrderVM(paymentOrderVO);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._OrderPrint, model);

            return View(MVC.Shared.Views._PrintHtml, (object)htmlContent);
        }

        [HttpPost]
        [EserviceAuthorize]
        public virtual FileResult DownloadPdfOrder(AuthRequestDO requestDO)
        {
            AisRequestDO aisRequestDO = ValidateAndGetAisRequestDO(requestDO);

            var paymentOrderVO = this.webRepository.GetPaymentOrderByGidAndUin(aisRequestDO.PaymentRequest.Gid, aisRequestDO.PaymentRequest.ApplicantUin);

            OrderVM model = new OrderVM(paymentOrderVO);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._OrderPrint, model);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "PaymentOrder" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        [NonAction]
        private AisRequestDO ValidateAndGetAisRequestDO(AuthRequestDO requestDO)
        {
            string dataJson = Encoding.UTF8.GetString(Convert.FromBase64String(requestDO.Data));
            JObject dataJObject = JObject.Parse(dataJson);

            string id = dataJObject.GetValue("id", StringComparison.OrdinalIgnoreCase).Value<string>();

            if (String.IsNullOrWhiteSpace(id))
            {
                throw new HttpException(400, "Invalid post data.");
            }

            EserviceClient eserviceClient = this.systemRepository.GetEserviceClientByClientId(requestDO.ClientId);
            PaymentRequest request = this.webRepository.GetPaymentRequestByIdentifier(id);

            if (request == null || request.EserviceClientId != eserviceClient.EserviceClientId)
            {
                throw new HttpException(400, "Payment request is invalid.");
            }

            return new AisRequestDO()
            {
                IsRequestValid = true,
                EserviceClient = eserviceClient,
                PaymentRequest = request
            };
        }
    }
}