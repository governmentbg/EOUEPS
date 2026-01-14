using ClosedXML.Excel;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Data.ViewObjects.Web;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using EPayments.Web.Auth;
using EPayments.Web.Common;
using EPayments.Web.DataObjects;
using EPayments.Web.Models.EserviceAdmin;
using EPayments.Web.Models.Shared;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace EPayments.Web.Controllers
{
    [EserviceAdminAuthorize]
    public partial class EserviceAdminController : BaseController
    {
        private IWebRepository webRepository;
        private ISystemRepository systemRepository;
        private IUnitOfWork unitOfWork;

        public EserviceAdminController(IWebRepository webRepository, ISystemRepository systemRepository, IUnitOfWork unitOfWork)
        {
            this.webRepository = webRepository;
            this.systemRepository = systemRepository;
            this.unitOfWork = unitOfWork;
        }

        [HttpGet]
        public virtual ActionResult PaymentRequests(EserviceAdminRequestSearchDO searchDO)
        {
            var eserviceAdminUser = this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.EserviceAdminUserId == this.CurrentUser.EserviceAdminId.Value)
                .Single();

            EserviceClient eserviceClient = this.unitOfWork.DbContext.Set<EserviceClient>()
                .Single(e => e.EserviceClientId == eserviceAdminUser.ReferringEserviceClientId.Value);

            PaymentRequestsVM model = new PaymentRequestsVM();
            model.EserviceClientName = eserviceClient.AisName;
            //model.EserviceObligationType = eserviceClient.ObligationType.Name;
            //model.EserviceObligationTypeCode = eserviceClient.ObligationType.PaymentTypeCode;

            model.RequestsPagingOptions = new PagingVM();

            model.SearchDO = searchDO;

            //Processed payments

            model.RequestsPagingOptions.CurrentPageIndex = searchDO.PrPage;
            model.RequestsPagingOptions.ControllerName = MVC.EserviceAdmin.Name;
            model.RequestsPagingOptions.ActionName = MVC.EserviceAdmin.ActionNames.PaymentRequests;
            model.RequestsPagingOptions.PageIndexParameterName = "prPage";
            model.RequestsPagingOptions.RouteValues = searchDO.ToRequestsRouteValues();

            model.RequestsPagingOptions.TotalItemCount = this.webRepository.CountEserviceAdminRequests(
                eserviceAdminUser.ReferringEserviceClientId.Value,
                searchDO.PrId,
                searchDO.PrRefNumber,
                searchDO.PrObligationType,
                searchDO.PrRefId,
                searchDO.PrPaymentType,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                searchDO.PrProvider,
                searchDO.PrReason,
                searchDO.PrStatus,
                searchDO.PrObligationStatus,
                searchDO.PrApplicantName,
                searchDO.PrApplicantUin);

            model.Requests = this.webRepository.GetEserviceAdminRequests(
                eserviceAdminUser.ReferringEserviceClientId.Value,
                searchDO.PrId,
                searchDO.PrRefNumber,
                searchDO.PrObligationType,
                searchDO.PrRefId,
                searchDO.PrPaymentType,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                searchDO.PrProvider,
                searchDO.PrReason,
                searchDO.PrStatus,
                searchDO.PrObligationStatus,
                searchDO.PrApplicantName,
                searchDO.PrApplicantUin,
                searchDO.PrSortBy,
                searchDO.PrSortDesc,
                searchDO.PrPage,
                AppSettings.EPaymentsWeb_MaxSearchResultsPerPage);

            model.PaymentRequestStatuses = this.unitOfWork.DbContext.Set<PaymentRequestStatusModel>().ToList();

            return View(model);
        }

        [HttpPost]
        public virtual ActionResult PaymentRequestsSearch(EserviceAdminRequestSearchDO searchDO)
        {
            TempData[TempDataKeys.SearchPerformed] = true;

            searchDO.PrPage = 1;
            //searchDO.Focus = Constants.ProcessedPaymentsFocusId;

            return RedirectToPaymentRequestsAction(searchDO);
        }

        [HttpGet]
        public virtual ActionResult PaymentRequestsSort(EserviceAdminRequestSearchDO searchDO)
        {
            return RedirectToPaymentRequestsAction(searchDO);
        }

        public virtual ActionResult PaymentRequestsExportAsPdf(EserviceAdminRequestSearchDO searchDO)
        {
            var eserviceAdminUser = this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.EserviceAdminUserId == this.CurrentUser.EserviceAdminId.Value)
                .Single();

            var requests = this.webRepository.GetEserviceAdminRequests(
                 eserviceAdminUser.ReferringEserviceClientId.Value,
                 searchDO.PrId,
                 searchDO.PrRefNumber,
                 searchDO.PrObligationType,
                 searchDO.PrRefId,
                 searchDO.PrPaymentType,
                 Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                 Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                 Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                 Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                 searchDO.PrProvider,
                 searchDO.PrReason,
                 searchDO.PrStatus,
                 searchDO.PrObligationStatus,
                 searchDO.PrApplicantName,
                 searchDO.PrApplicantUin,
                 searchDO.PrSortBy,
                 searchDO.PrSortDesc,
                 AppSettings.EPaymentsCommon_MaxPdfExportEntriesCount);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._PaymentRequestsPrintPdf, requests);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "Заявления" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        public virtual ActionResult PaymentRequestsExportAsExcel(EserviceAdminRequestSearchDO searchDO)
        {
            var eserviceAdminUser = this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.EserviceAdminUserId == this.CurrentUser.EserviceAdminId.Value)
                .Single();

            var requests = this.webRepository.GetEserviceAdminRequests(
                 eserviceAdminUser.ReferringEserviceClientId.Value,
                 searchDO.PrId,
                 searchDO.PrRefNumber,
                 searchDO.PrObligationType,
                 searchDO.PrRefId,
                 searchDO.PrPaymentType,
                 Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                 Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                 Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                 Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                 searchDO.PrProvider,
                 searchDO.PrReason,
                 searchDO.PrStatus,
                 searchDO.PrObligationStatus,
                 searchDO.PrApplicantName,
                 searchDO.PrApplicantUin,
                 searchDO.PrSortBy,
                 searchDO.PrSortDesc);

            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Транзакции");

            int row = 1;
            int col = 1;
            IXLCell cell;

            //Set column headers

            // A1
            cell = worksheet.Cell(row, col);
            cell.Value = "Номер на задължение";
            cell.Style.Font.Bold = true;

            // B1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Номер на документ (Референтен номер RNU)";
            cell.Style.Font.Bold = true;

            // C1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "refid - Номер на платежен документ МЕУ";
            cell.Style.Font.Bold = true;

            // D1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Вид плащане";
            cell.Style.Font.Bold = true;

            // E1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Основание за плащане";
            cell.Style.Font.Bold = true;

            if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
            {
                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Сума, BGN";
                cell.Style.Font.Bold = true;
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
            {
                if (CurrencyHelper.IsEuroTimePeriod)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Сума, EUR";
                    cell.Style.Font.Bold = true;

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Сума, BGN";
                    cell.Style.Font.Bold = true;
                }
                else
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Сума, BGN";
                    cell.Style.Font.Bold = true;

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Сума, EUR";
                    cell.Style.Font.Bold = true;
                }
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
            {
                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Сума, EUR";
                cell.Style.Font.Bold = true;
            }

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Задължено лице";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "ЕГН / ЕИК (Булстат)";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дата и час на плащане";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дата на изтичане на задължението";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Статус на плащане";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Статус на задължение";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Още пояснения";
            cell.Style.Font.Bold = true;

            // Move to the begining of row 2
            col = 1;
            row = 2;

            for (int i = 0; i < requests.Count; i++)
            {
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].PaymentRequestIdentifier);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].PaymentReferenceNumber);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].Refid);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].ObligationType);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].PaymentReason);

                if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetBgnValueFormated(requests[i].PaymentAmountRequest, requests[i].CreateDate));
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
                {
                    if (CurrencyHelper.IsEuroTimePeriod)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(requests[i].PaymentAmountRequest, requests[i].CreateDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(requests[i].PaymentAmountRequest, requests[i].CreateDate));
                    }
                    else
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(requests[i].PaymentAmountRequest, requests[i].CreateDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(requests[i].PaymentAmountRequest, requests[i].CreateDate));
                    }
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(requests[i].PaymentAmountRequest, requests[i].CreateDate));
                }

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].ApplicantName);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].ApplicantUin);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.DateTimeToBgFormatWithoutSecondsNotLocalTime(requests[i].TransactionDate));

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.DateTimeToBgFormatWithoutSecondsNotLocalTime(requests[i].ExpirationDate));

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.EnumToDescriptionString(requests[i].PaymentRequestStatusId));

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].ObligationStatusId != null ? Formatter.EnumToDescriptionString(requests[i].ObligationStatusId) : "Няма стойност");

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(requests[i].AdditionalInformation);

                col = 1;
                row++;
            }

            worksheet.Rows().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Rows().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

            worksheet.Column("A").AdjustToContents();
            worksheet.Column("B").AdjustToContents();
            worksheet.Column("C").AdjustToContents();
            worksheet.Column("D").AdjustToContents();
            worksheet.Column("E").AdjustToContents();
            worksheet.Column("F").AdjustToContents();
            worksheet.Column("G").AdjustToContents();
            worksheet.Column("H").AdjustToContents();
            worksheet.Column("I").AdjustToContents();
            worksheet.Column("J").AdjustToContents();
            worksheet.Column("K").AdjustToContents();
            worksheet.Column("L").AdjustToContents();
            worksheet.Column("M").AdjustToContents();
            worksheet.Column("N").AdjustToContents();

            MemoryStream excelStream = new MemoryStream();
            workbook.SaveAs(excelStream);
            excelStream.Position = 0;

            return File(excelStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Заявления.xlsx");
        }

        [HttpGet]
        public virtual ActionResult CreatePaymentRequest()
        {
            EserviceAdminUser eserviceAdminUser = this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.EserviceAdminUserId == this.CurrentUser.EserviceAdminId.Value)
                .Single();

            EserviceClient eserviceClient = this.unitOfWork.DbContext.Set<EserviceClient>()
                .Single(e => e.EserviceClientId == eserviceAdminUser.ReferringEserviceClientId.Value);

            PaymentRequestVM model = new PaymentRequestVM();
            //model.PaymentRequest.ObligationType = eserviceClient?.ObligationType?.Name ?? string.Empty;


            //PaymentRequestDO.ObligationTypesList = this.unitOfWork.DbContext.Set<ObligationType>().Where(o => o.ObligationTypeId == 1).ToList();
            PaymentRequestDO.ObligationTypesList.AddRange(this.unitOfWork.DbContext.Set<ObligationType>()
                                                                        .Where(o => o.AlgorithmId == 0)
                                                                        .Where(o => o.IsActive == true).ToList());
            return View(model);
        }

        [HttpPost]
        public virtual async Task<ActionResult> CreatePaymentRequest(PaymentRequestDO model)
        {
            string currency = CurrencyHelper.GetCurrencyCode();

            EserviceAdminUser eserviceAdminUser = this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.EserviceAdminUserId == this.CurrentUser.EserviceAdminId.Value)
                .Single();

            EserviceClient eserviceClient = this.unitOfWork.DbContext.Set<EserviceClient>()
                .Single(e => e.EserviceClientId == eserviceAdminUser.ReferringEserviceClientId.Value);

            if (!this.ModelState.IsValid)
            {
                PaymentRequestVM paymentRequest = new PaymentRequestVM();
                paymentRequest.PaymentRequest = model;
                //paymentRequest.PaymentRequest.ObligationType = eserviceClient?.ObligationType?.Name ?? string.Empty;

                return View(paymentRequest);
            }

            if (model.PaymentReferenceType == null || model.PaymentReferenceNumber.Length < 1)
            {
                model.PaymentReferenceType = " ";
            }

            int obligationTypeId = int.Parse(model.ObligationType);
            string body = Convert.ToBase64String(
                System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(new
                {
                    serviceProviderName = eserviceClient.AisName,
                    serviceProviderBank = eserviceClient.AccountBank,
                    serviceProviderBIC = eserviceClient.AccountBIC,
                    serviceProviderIBAN = eserviceClient.AccountIBAN,
                    currency = currency,
                    paymentAmount = model.PaymentAmount.ToString("#.00", System.Globalization.CultureInfo.InvariantCulture),
                    paymentReason = model.PaymentReason,
                    obligationType = model.ObligationType,
                    applicantUinTypeId = (int)model.ApplicantUinTypeId,
                    applicantUin = model.ApplicantUin,
                    applicantName = model.ApplicantName,
                    paymentReferenceNumber = model.PaymentReferenceNumber,
                    paymentReferenceDate = DateTime.Now.ToString("s"),
                    expirationDate = model.ExpirationDate.ToString("s"),
                    paymentReferenceType = int.Parse(model.PaymentReferenceType),
                    additionalInformation = model.AdditionalInformation,
                    paymentTypeCode = this.unitOfWork.DbContext.Set<ObligationType>().Where(o => o.ObligationTypeId == obligationTypeId).Single().PaymentTypeCode.Trim(),
                })));

            string signature = HmacRequestHelper.CalculateHmac(eserviceClient.SecretKey, body);

            Dictionary<string, string> keyValues = new Dictionary<string, string>()
            {
                { "clientId", eserviceClient.ClientId },
                { "hmac", signature },
                { "data", body }
            };

            string content = null;

            using (HttpClient client = new HttpClient())
            {
                System.Net.ServicePointManager.SecurityProtocol |=
                    System.Net.SecurityProtocolType.Tls12 |
                    System.Net.SecurityProtocolType.Tls11 |
                    System.Net.SecurityProtocolType.Tls;

                try
                {
                    HttpResponseMessage response = await client.PostAsync(AppSettings.EPaymentsWeb_CreatePaymentRequest, new FormUrlEncodedContent(keyValues));

                    if (response.StatusCode != System.Net.HttpStatusCode.OK)
                    {
                        PaymentRequestVM paymentRequest = new PaymentRequestVM();
                        ModelState.AddModelError(string.Empty, "Грешка при създаване на заявка за плащане");
                        paymentRequest.PaymentRequest = model;
                        paymentRequest.PaymentRequest.ObligationType = model.ObligationType;

                        return View(paymentRequest);
                    }

                    content = await response.Content.ReadAsStringAsync();
                }
                catch (Exception ex)
                {
                    ModelState.AddModelError(string.Empty, string.Format("Грешка при създаване на заявка за плащане:{0}{1}",
                        Environment.NewLine,
                        ex.Message));

                    PaymentRequestVM paymentRequest = new PaymentRequestVM();
                    paymentRequest.PaymentRequest = model;
                    paymentRequest.PaymentRequest.ObligationType = model.ObligationType;

                    return View(paymentRequest);
                }
            }

            PaymentRequestResponseDO paymentRequestResponseDO = JsonConvert.DeserializeObject<PaymentRequestResponseDO>(content);

            if (paymentRequestResponseDO.UnacceptedReceiptJson != null &&
                paymentRequestResponseDO.UnacceptedReceiptJson.Errors != null &&
                paymentRequestResponseDO.UnacceptedReceiptJson.Errors.Count > 0)
            {
                paymentRequestResponseDO.UnacceptedReceiptJson.Errors.ForEach(e => ModelState.AddModelError(string.Empty, e));

                PaymentRequestVM paymentRequest = new PaymentRequestVM();
                paymentRequest.PaymentRequest = model;
                paymentRequest.PaymentRequest.ObligationType = model.ObligationType;

                return View(paymentRequest);
            }

            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests);
        }

        [HttpPost]
        public virtual async Task<ActionResult> SuspendRequest(Guid gid)
        {
            EserviceAdminUser eserviceAdminUser = this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.EserviceAdminUserId == this.CurrentUser.EserviceAdminId.Value)
                .Single();

            EserviceClient eserviceClient = this.unitOfWork.DbContext.Set<EserviceClient>()
                .Single(e => e.EserviceClientId == eserviceAdminUser.ReferringEserviceClientId.Value);

            PaymentRequest paymentRequest = this.unitOfWork.DbContext.Set<PaymentRequest>()
                .SingleOrDefault(pr => pr.Gid == gid);

            if (paymentRequest == null)
            {
                TempData[TempDataKeys.IsPaymentRequestError] = false;
                TempData[TempDataKeys.PaymentRequestError] = "Заявката за плащане не е намерена.";

                return this.RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
            }

            if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Pending || paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Canceled)
            {
                return await this.UpdatePayment(eserviceClient, AppSettings.EPaymentsWeb_SuspendPaymentRequest, new
                {
                    id = paymentRequest.PaymentRequestIdentifier
                });
            }

            TempData[TempDataKeys.IsPaymentRequestError] = false;
            TempData[TempDataKeys.PaymentRequestError] = string.Format("Можете да променяте заявки за плащане които са в статус {0} и {1}.",
                PaymentRequestStatus.Pending.GetDescription(),
                PaymentRequestStatus.Canceled.GetDescription());

            return this.RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
        }

        [HttpPost]
        public virtual async Task<ActionResult> SetStatusPaid(Guid gid, string paymentMethod)
        {
            EserviceAdminUser eserviceAdminUser = this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.EserviceAdminUserId == this.CurrentUser.EserviceAdminId.Value)
                .Single();

            EserviceClient eserviceClient = this.unitOfWork.DbContext.Set<EserviceClient>()
                .Single(e => e.EserviceClientId == eserviceAdminUser.ReferringEserviceClientId.Value);

            PaymentRequest paymentRequest = this.unitOfWork.DbContext.Set<PaymentRequest>()
                .SingleOrDefault(pr => pr.Gid == gid);

            if (paymentRequest == null)
            {
                TempData[TempDataKeys.IsPaymentRequestError] = false;
                TempData[TempDataKeys.PaymentRequestError] = "Заявката за плащане не е намерена.";

                return this.RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
            }

            if (!Enum.TryParse(paymentMethod, out PaidStatusPaymentMethod value))
            {
                TempData[TempDataKeys.IsPaymentRequestError] = false;
                TempData[TempDataKeys.PaymentRequestError] = "Невалиден начин на плащане.";

                return this.RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
            }

            if (paymentRequest.PaymentRequestStatusId == PaymentRequestStatus.Pending)
            {
                return await this.UpdatePayment(eserviceClient, AppSettings.EPaymentsWeb_SetStatusPaidPaymentRequest,
                    new
                    {
                        id = paymentRequest.PaymentRequestIdentifier,
                        paymentMethod = paymentMethod,
                        paymentDescription = string.Empty
                    });
            }

            TempData[TempDataKeys.IsPaymentRequestError] = false;
            TempData[TempDataKeys.PaymentRequestError] = string.Format("Можете да променяте заявки за плащане които са в статус {0}.",
                PaymentRequestStatus.Pending.GetDescription());

            return this.RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
        }

        [NonAction]
        private async Task<ActionResult> UpdatePayment(EserviceClient eserviceClient, string address, object values)
        {
            string body = Convert.ToBase64String(
                System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(values)));

            string signature = HmacRequestHelper.CalculateHmac(eserviceClient.SecretKey, body);

            Dictionary<string, string> keyValues = new Dictionary<string, string>()
                {
                    { "clientId", eserviceClient.ClientId },
                    { "hmac", signature },
                    { "data", body }
                };

            using (HttpClient client = new HttpClient())
            {
                System.Net.ServicePointManager.SecurityProtocol |=
                    System.Net.SecurityProtocolType.Tls12 |
                    System.Net.SecurityProtocolType.Tls11 |
                    System.Net.SecurityProtocolType.Tls;

                try
                {
                    HttpResponseMessage response = await client.PostAsync(address, new FormUrlEncodedContent(keyValues));

                    if (response.StatusCode != System.Net.HttpStatusCode.OK)
                    {
                        TempData[TempDataKeys.IsPaymentRequestError] = false;
                        TempData[TempDataKeys.PaymentRequestError] = "Заявката за плащане не беше обновена.";

                        return this.RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
                    }
                }
                catch (Exception ex)
                {
                    TempData[TempDataKeys.IsPaymentRequestError] = false;
                    TempData[TempDataKeys.PaymentRequestError] = "Грешка при обновяване на заявка за плащане. Заявката за плащане не беше обновена. ";

                    return this.RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
                }
            }

            TempData[TempDataKeys.IsPaymentRequestError] = true;
            TempData[TempDataKeys.PaymentRequestError] = "Заявката за плащане беше обновена.";

            return this.RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);
        }

        [NonAction]
        private ActionResult RedirectToPaymentRequestsAction(EserviceAdminRequestSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name,
                new
                {
                    @prId = searchDO.PrId,
                    @prRefNumber = searchDO.PrRefNumber,
                    @prObligationType = searchDO.PrObligationType,
                    @prRefId = searchDO.PrRefId,
                    @prPaymentType = searchDO.PrPaymentType,
                    @prDateFrom = searchDO.PrDateFrom,
                    @prDateTo = searchDO.PrDateTo,
                    @prAmountFrom = searchDO.PrAmountFrom,
                    @prAmountTo = searchDO.PrAmountTo,
                    @prProvider = searchDO.PrProvider,
                    @prReason = searchDO.PrReason,
                    @prStatus = searchDO.PrStatus,
                    @prObligationStatus = searchDO.PrObligationStatus,

                    @prApplicantName = searchDO.PrApplicantName,
                    @prApplicantUin = searchDO.PrApplicantUin,

                    @prPage = searchDO.PrPage,
                    @prSortBy = searchDO.PrSortBy,
                    @prSortDesc = searchDO.PrSortDesc,

                    @focus = searchDO.Focus,
                });
        }

        [HttpPost]
        public virtual ActionResult PrintPaymentRequests(EserviceAdminRequestSearchDO searchDO, bool printAllResults = false)
        {
            var eserviceAdminUser = this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .Where(e => e.EserviceAdminUserId == this.CurrentUser.EserviceAdminId.Value)
                .Single();

            PrintPaymentRequestsVM model = new PrintPaymentRequestsVM();

            model.PrintAllResults = printAllResults;

            model.Requests = this.webRepository.GetEserviceAdminRequests(
                eserviceAdminUser.ReferringEserviceClientId.Value,
                searchDO.PrId,
                searchDO.PrRefNumber,
                searchDO.PrPaymentType,
                searchDO.PrObligationType,
                searchDO.PrRefId,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                searchDO.PrProvider,
                searchDO.PrReason,
                searchDO.PrStatus,
                searchDO.PrObligationStatus,
                searchDO.PrApplicantName,
                searchDO.PrApplicantUin,
                searchDO.PrSortBy,
                searchDO.PrSortDesc,
                !printAllResults ? searchDO.PrPage : (int?)null,
                !printAllResults ? AppSettings.EPaymentsWeb_MaxSearchResultsPerPage : (int?)null);

            return View(model);
        }

        [HttpGet]
        public virtual ActionResult TransactionList(TransactionListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //List<EserviceBankAccount> authorizedBankAccounts = this.webRepository.GetEserviceBankAccountsByAdminUserId(this.CurrentUser.EserviceAdminId.Value);

            //TransactionListVM model = new TransactionListVM();

            ////set BankAndIbanDictionary

            //model.BankAndIbanDictionary = new Dictionary<string, List<string>>();

            //foreach (EserviceBankAccount authorizedBankAccount in authorizedBankAccounts)
            //{
            //    if (!model.BankAndIbanDictionary.ContainsKey(authorizedBankAccount.Bank.Name))
            //    {
            //        model.BankAndIbanDictionary.Add(authorizedBankAccount.Bank.Name, new List<string>());
            //    }

            //    model.BankAndIbanDictionary[authorizedBankAccount.Bank.Name].Add(authorizedBankAccount.Iban);
            //}

            ////set BankAccounts

            //model.BankAccounts = authorizedBankAccounts.Select(e => new SelectListItem
            //{
            //    Value = e.EserviceBankAccountId.ToString(),
            //    Text = e.Iban
            //}).ToList();

            //model.TransactionRecordsPagingOptions = new PagingVM();

            //model.SearchDO = searchDO;

            //if (String.IsNullOrWhiteSpace(Request.QueryString.ToString()))
            //{
            //    if (model.BankAccounts.Any())
            //    {
            //        model.SearchDO.EserviceBankAccountId = int.Parse(model.BankAccounts.First().Value);
            //    }
            //    model.SearchDO.TransactionRecordRefStatus = -1;
            //}

            //model.TransactionRecordsPagingOptions.CurrentPageIndex = searchDO.Page;
            //model.TransactionRecordsPagingOptions.ControllerName = MVC.EserviceAdmin.Name;
            //model.TransactionRecordsPagingOptions.ActionName = MVC.EserviceAdmin.ActionNames.TransactionList;
            //model.TransactionRecordsPagingOptions.PageIndexParameterName = "page";
            //model.TransactionRecordsPagingOptions.RouteValues = searchDO.ToRouteValues();

            //if (searchDO.EserviceBankAccountId.HasValue)
            //{
            //    if (authorizedBankAccounts.Any(e => e.EserviceBankAccountId == searchDO.EserviceBankAccountId.Value))
            //    {
            //        Tuple <int, decimal> countAndTotalAmount = this.webRepository.CountAndGetTotalAmountForTransactionRecords(
            //            searchDO.EserviceBankAccountId.Value,
            //            Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.TransactionAccountingDateFrom)),
            //            Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.TransactionAccountingDateTo)),
            //            Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.TransactionAmountFrom),
            //            Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.TransactionAmountTo),
            //            searchDO.InfoDocumentNumber,
            //            Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.InfoDocumentDateFrom)),
            //            Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.InfoDocumentDateTo)),
            //            searchDO.InfoSenderIban,
            //            searchDO.InfoSenderName,
            //            searchDO.InfoDebtorName,
            //            searchDO.InfoDebtorBulstatEgnLnch,
            //            searchDO.InfoPaymentReason,
            //            searchDO.InfoAC1AuthorizationCode,
            //            searchDO.TransactionRecordPaymentMethod,
            //            searchDO.TransactionRecordRefStatus);

            //        model.TransactionRecordsPagingOptions.TotalItemCount = countAndTotalAmount.Item1;
            //        model.TotalAmount = countAndTotalAmount.Item2;

            //        model.TransactionRecords = this.webRepository.GetTransactionRecords(
            //            searchDO.EserviceBankAccountId.Value,
            //            Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.TransactionAccountingDateFrom)),
            //            Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.TransactionAccountingDateTo)),
            //            Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.TransactionAmountFrom),
            //            Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.TransactionAmountTo),
            //            searchDO.InfoDocumentNumber,
            //            Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.InfoDocumentDateFrom)),
            //            Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.InfoDocumentDateTo)),
            //            searchDO.InfoSenderIban,
            //            searchDO.InfoSenderName,
            //            searchDO.InfoDebtorName,
            //            searchDO.InfoDebtorBulstatEgnLnch,
            //            searchDO.InfoPaymentReason,
            //            searchDO.InfoAC1AuthorizationCode,
            //            searchDO.TransactionRecordPaymentMethod,
            //            searchDO.TransactionRecordRefStatus,
            //            searchDO.SortBy,
            //            searchDO.SortDesc,
            //            searchDO.Page,
            //            AppSettings.EPaymentsWeb_MaxTransactionResultsPerPage);
            //    }
            //}

            //return View(model);
        }

        [HttpPost]
        public virtual ActionResult TransactionListSearch(TransactionListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //searchDO.Page = 1;

            //return RedirectToTransactionListAction(searchDO);
        }

        [HttpGet]
        public virtual ActionResult TransactionListSort(TransactionListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //return RedirectToTransactionListAction(searchDO);
        }

        [HttpGet]
        public virtual ActionResult TransactionDetails(int id)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //TransactionDetailsVM model = new TransactionDetailsVM();

            //TransactionRecord transactionRecord = this.webRepository.GetTransactionRecordByTransactionRecordId(id);

            //if (transactionRecord != null)
            //{
            //    List<EserviceBankAccount> authorizedBankAccounts = this.webRepository.GetEserviceBankAccountsByAdminUserId(this.CurrentUser.EserviceAdminId.Value);

            //    if (authorizedBankAccounts.Any(e => e.EserviceBankAccountId == transactionRecord.TransactionFile.EserviceBankAccountId))
            //    {
            //        model.SetFields(transactionRecord);
            //    }
            //}

            //return PartialView(model);
        }

        [HttpGet]
        public virtual ActionResult TransactionReferencedPayment(Guid id, bool showDetailsForm = false)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //List<EserviceBankAccount> authorizedBankAccounts = this.webRepository.GetEserviceBankAccountsByAdminUserId(this.CurrentUser.EserviceAdminId.Value);

            //var paymentOrderVO = this.webRepository.GetPaymentOrderByGid(id);

            //if (!authorizedBankAccounts.Any(e => e.Iban == paymentOrderVO.IBAN))
            //    throw new Exception("Unauthorized access");

            //OrderVM model = new OrderVM(paymentOrderVO, null, true);
            //model.ShowDetailsForm = true;

            //return PartialView(MVC.Shared.Views._Order, model);
        }

        [HttpGet]
        public virtual ActionResult PrintOrder(Guid id)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //List<EserviceBankAccount> authorizedBankAccounts = this.webRepository.GetEserviceBankAccountsByAdminUserId(this.CurrentUser.EserviceAdminId.Value);

            //var paymentOrderVO = this.webRepository.GetPaymentOrderByGid(id);

            //if (!authorizedBankAccounts.Any(e => e.Iban == paymentOrderVO.IBAN))
            //    throw new Exception("Unauthorized access");

            //OrderVM model = new OrderVM(paymentOrderVO, null, true);

            //string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._OrderPrint, model);

            //return View(MVC.Shared.Views._PrintHtml, (object)htmlContent);
        }

        [HttpGet]
        public virtual FileResult DownloadPdfOrder(Guid id)
        {
            return null;

            //List<EserviceBankAccount> authorizedBankAccounts = this.webRepository.GetEserviceBankAccountsByAdminUserId(this.CurrentUser.EserviceAdminId.Value);

            //var paymentOrderVO = this.webRepository.GetPaymentOrderByGid(id);

            //if (!authorizedBankAccounts.Any(e => e.Iban == paymentOrderVO.IBAN))
            //    throw new Exception("Unauthorized access");

            //OrderVM model = new OrderVM(paymentOrderVO, null, true);

            //string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._OrderPrint, model);

            //byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            //string fileName = "PaymentOrder" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            //return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        [HttpGet]
        public virtual ActionResult EditSettings()
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //EserviceAdminUser eserviceAdminUser = this.systemRepository.GetActiveEserviceAdminUserById(this.CurrentUser.EserviceAdminId.Value);

            //EditSettingsVM model = new EditSettingsVM();

            //model.Email = eserviceAdminUser.Email;
            //model.InsufficientAmountNotifications = eserviceAdminUser.InsufficientAmountNotifications;
            //model.OverpaidAmountNotifications = eserviceAdminUser.OverpaidAmountNotifications;
            //model.ReferencedNotifications = eserviceAdminUser.ReferencedNotifications;
            //model.NotReferencedNotifications = eserviceAdminUser.NotReferencedNotifications;

            //return View(model);
        }

        [HttpPost]
        public virtual ActionResult EditSettings(EditSettingsVM model)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //if ((model.InsufficientAmountNotifications || model.OverpaidAmountNotifications || model.ReferencedNotifications || model.NotReferencedNotifications) &&
            //    String.IsNullOrWhiteSpace(model.Email))
            //{
            //    ModelState.AddModelError("Email", "Полето „Електронна поща“ е задължително.");
            //}

            //if (!ModelState.IsValid)
            //{
            //    return View(model);
            //}

            //EserviceAdminUser eserviceAdminUser = this.systemRepository.GetActiveEserviceAdminUserById(this.CurrentUser.EserviceAdminId.Value);

            //eserviceAdminUser.Email = !String.IsNullOrWhiteSpace(model.Email) ? model.Email : null;
            //eserviceAdminUser.InsufficientAmountNotifications = model.InsufficientAmountNotifications;
            //eserviceAdminUser.OverpaidAmountNotifications = model.OverpaidAmountNotifications;
            //eserviceAdminUser.ReferencedNotifications = model.ReferencedNotifications;
            //eserviceAdminUser.NotReferencedNotifications = model.NotReferencedNotifications;

            //this.unitOfWork.Save();

            //return RedirectToAction(MVC.EserviceAdmin.ActionNames.Settings, MVC.EserviceAdmin.Name);
        }

        [HttpGet]
        public virtual ActionResult Settings(RequestAccessListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //EserviceAdminUser eserviceAdminUser = this.systemRepository.GetActiveEserviceAdminUserById(this.CurrentUser.EserviceAdminId.Value);

            //SettingsVM model = new SettingsVM();

            //model.Email = eserviceAdminUser.Email;
            //model.InsufficientAmountNotifications = eserviceAdminUser.InsufficientAmountNotifications;
            //model.OverpaidAmountNotifications = eserviceAdminUser.OverpaidAmountNotifications;
            //model.ReferencedNotifications = eserviceAdminUser.ReferencedNotifications;
            //model.NotReferencedNotifications = eserviceAdminUser.NotReferencedNotifications;

            //return View(model);
        }

        [HttpGet]
        public virtual ActionResult TransactionListExportExcel(TransactionListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //Tuple<string, List<TransactionRecordVO>> ibanAndRecords = GetIbanAndTransactionRecords(searchDO);

            //string iban = ibanAndRecords.Item1;
            //IList<TransactionRecordVO> records = ibanAndRecords.Item2;

            //XLWorkbook workbook = new XLWorkbook();
            //var worksheet = workbook.Worksheets.Add("Транзакции");

            ////Set column headers
            //worksheet.Cell("A1").Value = "Дата";
            //worksheet.Cell("B1").Value = "Сума";
            //worksheet.Cell("C1").Value = "Документ";
            //worksheet.Cell("D1").Value = "Задължено лице";
            //worksheet.Cell("E1").Value = "Основание";
            //worksheet.Cell("F1").Value = "Наредител";
            //worksheet.Cell("G1").Value = "Плащане";
            //worksheet.Cell("H1").Value = "Статус";

            //worksheet.Cell("A1").Style.Font.Bold = true;
            //worksheet.Cell("B1").Style.Font.Bold = true;
            //worksheet.Cell("C1").Style.Font.Bold = true;
            //worksheet.Cell("D1").Style.Font.Bold = true;
            //worksheet.Cell("E1").Style.Font.Bold = true;
            //worksheet.Cell("F1").Style.Font.Bold = true;
            //worksheet.Cell("H1").Style.Font.Bold = true; worksheet.Cell("G1").Style.Font.Bold = true;


            //for (int i = 0; i < records.Count; i++)
            //{
            //    if (records[i].TransactionAccountingDate.HasValue)
            //    {
            //        worksheet.Cell(String.Format("A{0}", i + 2)).SetValue<string>(Formatter.DateToBgFormatWithoutYearSuffix(records[i].TransactionAccountingDate));

            //        //worksheet.Cell(String.Format("A{0}", i + 2)).Value = Formatter.DateToBgFormatWithoutYearSuffix(records[i].TransactionAccountingDate);
            //    }

            //    if (records[i].TransactionAmount.HasValue)
            //    {
            //        worksheet.Cell(String.Format("B{0}", i + 2)).Value = Formatter.DecimalToTwoDecimalPlacesFormat(records[i].TransactionAmount) + " лв.";
            //    }

            //    if (!String.IsNullOrWhiteSpace(records[i].InfoDocumentNumber) || records[i].InfoDocumentDate.HasValue)
            //    {
            //        string cellValue = null;

            //        if (!String.IsNullOrWhiteSpace(records[i].InfoDocumentNumber) && !records[i].InfoDocumentDate.HasValue)
            //        {
            //            cellValue = records[i].InfoDocumentNumber;
            //        }
            //        else if (String.IsNullOrWhiteSpace(records[i].InfoDocumentNumber) && records[i].InfoDocumentDate.HasValue)
            //        {
            //            cellValue = Formatter.DateToBgFormatWithoutYearSuffix(records[i].InfoDocumentDate.Value);
            //        }
            //        else if (!String.IsNullOrWhiteSpace(records[i].InfoDocumentNumber) && records[i].InfoDocumentDate.HasValue)
            //        {
            //            cellValue = String.Format("{0}{1}{2}",
            //                records[i].InfoDocumentNumber,
            //                Environment.NewLine,
            //                Formatter.DateToBgFormatWithoutYearSuffix(records[i].InfoDocumentDate.Value));
            //        }

            //        worksheet.Cell(String.Format("C{0}", i + 2)).Value = cellValue;
            //    }

            //    if (!String.IsNullOrWhiteSpace(records[i].InfoDebtorBulstatEgnLnch) || !String.IsNullOrWhiteSpace(records[i].InfoDebtorName))
            //    {
            //        string cellValue = null;

            //        if (!String.IsNullOrWhiteSpace(records[i].InfoDebtorBulstatEgnLnch) && String.IsNullOrWhiteSpace(records[i].InfoDebtorName))
            //        {
            //            cellValue = records[i].InfoDebtorBulstatEgnLnch;
            //        }
            //        else if (String.IsNullOrWhiteSpace(records[i].InfoDebtorBulstatEgnLnch) && !String.IsNullOrWhiteSpace(records[i].InfoDebtorName))
            //        {
            //            cellValue = records[i].InfoDebtorName;
            //        }
            //        else if (!String.IsNullOrWhiteSpace(records[i].InfoDebtorBulstatEgnLnch) && !String.IsNullOrWhiteSpace(records[i].InfoDebtorName))
            //        {
            //            cellValue = String.Format("{0}{1}{2}",
            //                records[i].InfoDebtorBulstatEgnLnch,
            //                Environment.NewLine,
            //                records[i].InfoDebtorName);
            //        }

            //        worksheet.Cell(String.Format("D{0}", i + 2)).Value = cellValue;
            //    }

            //    if (!String.IsNullOrWhiteSpace(records[i].InfoPaymentReason))
            //    {
            //        worksheet.Cell(String.Format("E{0}", i + 2)).Value = records[i].InfoPaymentReason;
            //    }

            //    if (!String.IsNullOrWhiteSpace(records[i].InfoSenderIban) || !String.IsNullOrWhiteSpace(records[i].InfoSenderName))
            //    {
            //        string cellValue = null;

            //        if (!String.IsNullOrWhiteSpace(records[i].InfoSenderIban) && String.IsNullOrWhiteSpace(records[i].InfoSenderName))
            //        {
            //            cellValue = records[i].InfoSenderIban;
            //        }
            //        else if (String.IsNullOrWhiteSpace(records[i].InfoSenderIban) && !String.IsNullOrWhiteSpace(records[i].InfoSenderName))
            //        {
            //            cellValue = records[i].InfoSenderName;
            //        }
            //        else if (!String.IsNullOrWhiteSpace(records[i].InfoSenderIban) && !String.IsNullOrWhiteSpace(records[i].InfoSenderName))
            //        {
            //            cellValue = String.Format("{0}{1}{2}",
            //                records[i].InfoSenderIban,
            //                Environment.NewLine,
            //                records[i].InfoSenderName);
            //        }

            //        worksheet.Cell(String.Format("F{0}", i + 2)).Value = cellValue;
            //    }

            //    worksheet.Cell(String.Format("G{0}", i + 2)).Value =
            //        records[i].TransactionRecordPaymentMethodId == TransactionRecordPaymentMethod.BankOrder ? "По банка" :
            //        (records[i].TransactionRecordPaymentMethodId == TransactionRecordPaymentMethod.POS ? "POS" :
            //        (records[i].TransactionRecordPaymentMethodId == TransactionRecordPaymentMethod.VPOS ? "VPOS" : String.Empty));

            //    worksheet.Cell(String.Format("H{0}", i + 2)).Value = Formatter.EnumToDescriptionString(records[i].TransactionRecordRefStatusId);
            //}

            ////align row content

            //worksheet.Rows().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            //worksheet.Rows().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

            ////adjust cells content

            //worksheet.Column("A").AdjustToContents();
            //worksheet.Column("B").AdjustToContents();
            //worksheet.Column("C").AdjustToContents();
            //worksheet.Column("D").AdjustToContents();
            //worksheet.Column("E").AdjustToContents();
            //worksheet.Column("F").AdjustToContents();
            //worksheet.Column("G").AdjustToContents();
            //worksheet.Column("H").AdjustToContents();

            //// Send the file
            //MemoryStream excelStream = new MemoryStream();
            //workbook.SaveAs(excelStream);
            //excelStream.Position = 0;

            //return File(excelStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", String.Format("Transactions {0}.xlsx", iban));
        }

        [HttpGet]
        public virtual ActionResult TransactionListExportHtml(TransactionListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //Tuple<string, List<TransactionRecordVO>> ibanAndRecords = GetIbanAndTransactionRecords(searchDO);

            //TransactionListPrintHtmlVM model = new TransactionListPrintHtmlVM();
            //model.PrintResult = false;
            //model.Records = ibanAndRecords.Item2;

            ////string Result = Engine.Razor.RunCompile(template, "templateKey", null, new { Name = "World" })
            //string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Views", "EserviceAdmin", "TransactionListPrintHtml.cshtml");
            //LoadedTemplateSource templateSource = new LoadedTemplateSource(System.IO.File.ReadAllText(templatePath), "TransactionListPrintHtml.cshtml");
            //string htmlContent = Engine.Razor.RunCompile(templateSource, "TransactionListPrintHtml.cshtml", typeof(TransactionListPrintHtmlVM), model);

            //// Send the file
            //MemoryStream htmlStream = new MemoryStream();
            //StreamWriter writer = new StreamWriter(htmlStream);
            //writer.Write(htmlContent);
            //writer.Flush();
            //htmlStream.Position = 0;

            //return File(htmlStream, "text/html", String.Format("Transactions {0}.html", ibanAndRecords.Item1));
        }

        [HttpGet]
        public virtual ActionResult TransactionListExportXml(TransactionListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //Tuple<string, List<TransactionRecordVO>> ibanAndRecords = GetIbanAndTransactionRecords(searchDO);

            //var xmlSerializer = new XmlSerializer(typeof(List<TransactionRecordVO>));
            //MemoryStream xmlStream = new MemoryStream();
            //xmlSerializer.Serialize(xmlStream, ibanAndRecords.Item2);
            //xmlStream.Position = 0;

            //return File(xmlStream, "text/xml", String.Format("Transactions {0}.xml", ibanAndRecords.Item1));
        }

        [HttpGet]
        public virtual ActionResult TransactionListPrintHtml(TransactionListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.PaymentRequests, MVC.EserviceAdmin.Name);

            //Tuple<string, List<TransactionRecordVO>> ibanAndRecords = GetIbanAndTransactionRecords(searchDO);

            //TransactionListPrintHtmlVM model = new TransactionListPrintHtmlVM();
            //model.PrintResult = true;
            //model.Records = ibanAndRecords.Item2;

            //return View(model);
        }

        [NonAction]
        private Tuple<string, List<TransactionRecordVO>> GetIbanAndTransactionRecords(TransactionListSearchDO searchDO)
        {
            List<EserviceBankAccount> authorizedBankAccounts = this.webRepository.GetEserviceBankAccountsByAdminUserId(this.CurrentUser.EserviceAdminId.Value);

            if (searchDO.EserviceBankAccountId.HasValue)
            {
                if (authorizedBankAccounts.Any(e => e.EserviceBankAccountId == searchDO.EserviceBankAccountId.Value))
                {
                    string iban = authorizedBankAccounts.First(e => e.EserviceBankAccountId == searchDO.EserviceBankAccountId.Value).Iban;

                    IList<TransactionRecordVO> records = this.webRepository.GetTransactionRecords(
                        searchDO.EserviceBankAccountId.Value,
                        Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.TransactionAccountingDateFrom)),
                        Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.TransactionAccountingDateTo)),
                        Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.TransactionAmountFrom),
                        Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.TransactionAmountTo),
                        searchDO.InfoDocumentNumber,
                        Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.InfoDocumentDateFrom)),
                        Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.InfoDocumentDateTo)),
                        searchDO.InfoSenderIban,
                        searchDO.InfoSenderName,
                        searchDO.InfoDebtorName,
                        searchDO.InfoDebtorBulstatEgnLnch,
                        searchDO.InfoPaymentReason,
                        searchDO.InfoAC1AuthorizationCode,
                        searchDO.TransactionRecordPaymentMethod,
                        searchDO.TransactionRecordRefStatus,
                        searchDO.SortBy,
                        searchDO.SortDesc);

                    return new Tuple<string, List<TransactionRecordVO>>(iban, records.ToList());
                }
            }

            throw new Exception("Unauthorized access");
        }

        [NonAction]
        private ActionResult RedirectToTransactionListAction(TransactionListSearchDO searchDO)
        {
            return RedirectToAction(MVC.EserviceAdmin.ActionNames.TransactionList, MVC.EserviceAdmin.Name,
                new
                {
                    @eserviceBankAccountId = searchDO.EserviceBankAccountId,
                    @transactionAccountingDateFrom = searchDO.TransactionAccountingDateFrom,
                    @transactionAccountingDateTo = searchDO.TransactionAccountingDateTo,
                    @transactionAmountFrom = searchDO.TransactionAmountFrom,
                    @transactionAmountTo = searchDO.TransactionAmountTo,
                    @infoDocumentNumber = searchDO.InfoDocumentNumber,
                    @infoDocumentDateFrom = searchDO.InfoDocumentDateFrom,
                    @infoDocumentDateTo = searchDO.InfoDocumentDateTo,
                    @infoSenderIban = searchDO.InfoSenderIban,
                    @infoSenderName = searchDO.InfoSenderName,
                    @infoDebtorName = searchDO.InfoDebtorName,
                    @infoDebtorBulstatEgnLnch = searchDO.InfoDebtorBulstatEgnLnch,
                    @infoPaymentReason = searchDO.InfoPaymentReason,
                    @infoAC1AuthorizationCode = searchDO.InfoAC1AuthorizationCode,
                    @transactionRecordPaymentMethod = searchDO.TransactionRecordPaymentMethod,
                    @transactionRecordRefStatus = searchDO.TransactionRecordRefStatus,

                    @page = searchDO.Page,
                    @sortBy = searchDO.SortBy,
                    @sortDesc = searchDO.SortDesc,
                });
        }
    }
}