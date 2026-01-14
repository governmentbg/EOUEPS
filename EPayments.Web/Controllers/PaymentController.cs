using ClosedXML.Excel;
using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Data.ViewObjects.Web;
using EPayments.Model.DataObjects;
using EPayments.Model.DataObjects.EmailTemplateContext;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using EPayments.Web.Auth;
using EPayments.Web.Common;
using EPayments.Web.DataObjects;
using EPayments.Web.Models.Payments;
using EPayments.Web.Models.Shared;
using log4net;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace EPayments.Web.Controllers
{
    [Authorize]
    public partial class PaymentController : BaseController
    {
        private IWebRepository webRepository;
        private ICommonRepository commonRepository;
        private ISystemRepository systemRepository;
        private readonly ILog logger;
        private IUnitOfWork unitOfWork;

        public PaymentController(IUnitOfWork unitOfWork, IWebRepository webRepository, ICommonRepository commonRepository, ISystemRepository systemRepository)
        {
            this.unitOfWork = unitOfWork;
            this.webRepository = webRepository;
            this.commonRepository = commonRepository;
            this.systemRepository = systemRepository;
            this.logger = LogManager.GetLogger(typeof(PaymentController));
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult List(ListSearchDO searchDO)
        {
            ListVM model = new ListVM();

            model.PendingPaymentsPagingOptions = new PagingVM();
            model.ProcessedPaymentsPagingOptions = new PagingVM();

            model.SearchDO = searchDO;

            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                //Pending payments

                model.PendingPaymentsPagingOptions.CurrentPageIndex = searchDO.PPage;
                model.PendingPaymentsPagingOptions.ControllerName = MVC.Payment.Name;
                model.PendingPaymentsPagingOptions.ActionName = MVC.Payment.ActionNames.List;
                model.PendingPaymentsPagingOptions.PageIndexParameterName = "pPage";
                model.PendingPaymentsPagingOptions.RouteValues = searchDO.ToPendingRequestsRouteValues();

                // NonMdt => PayOrder IS NULL
                List<PendingRequestVO> pendingPaymentsNonMdt = this.webRepository.GetPendingRequestsByUin(
                    this.CurrentUser.Uin,
                    searchDO.PSortBy,
                    searchDO.PSortDesc)
                    .Where(p => p.PayOrder == null).ToList();

                // Mdt => PayOrder IS NOT NULL and we need vakid JSON in AdditionalInfo
                List<PendingRequestVO> pendingPaymentsJsonCheck = this.webRepository.GetPendingRequestsByUin(
                    this.CurrentUser.Uin,
                    searchDO.PSortBy,
                    searchDO.PSortDesc)
                    .Where(p => p.PayOrder != null).ToList();

                List<PendingRequestVO> pendingPaymentsJsonValidated = new List<PendingRequestVO>();

                foreach (var pp in pendingPaymentsJsonCheck)
                {
                    try
                    {
                        var jsonParsed = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(pp.AdditionalInfo));
                        pendingPaymentsJsonValidated.Add(pp);
                    }
                    catch
                    {
                        this.logger.Error($"Incorrectly formed JSON: ({pp.AdditionalInfo}), ApplicantUin: {this.CurrentUser.Uin}, PaymentRequestIdentifier: {pp.PaymentRequestIdentifier}");
                    }
                }

                List<PendingRequestVO> allPendingPayments = new List<PendingRequestVO>();

                allPendingPayments.AddRange(pendingPaymentsNonMdt);

                allPendingPayments.AddRange(pendingPaymentsJsonValidated.
                                            OrderBy(p => p.EserviceClientId).
                                            ThenBy(p => p.ObligationTypeId).
                                            ThenBy(p => p.PayOrder).
                                            ThenBy(p => JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(p.AdditionalInfo)).PartidaNo).
                                            ThenBy(p => JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(p.AdditionalInfo)).InstNo)
                                        );

                model.PendingPaymentsPagingOptions.TotalItemCount = allPendingPayments.Count();

                model.PendingPayments = allPendingPayments.ToList();



                //Processed payments

                model.ProcessedPaymentsPagingOptions.CurrentPageIndex = searchDO.PrPage;
                model.ProcessedPaymentsPagingOptions.ControllerName = MVC.Payment.Name;
                model.ProcessedPaymentsPagingOptions.ActionName = MVC.Payment.ActionNames.List;
                model.ProcessedPaymentsPagingOptions.PageIndexParameterName = "prPage";
                model.ProcessedPaymentsPagingOptions.RouteValues = searchDO.ToProcessedRequestsRouteValues();

                model.ProcessedPaymentsPagingOptions.TotalItemCount = this.webRepository.CountProcessedRequestsByUin(
                    this.CurrentUser.Uin,
                    searchDO.PrId,
                    Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                    Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                    searchDO.PrProvider,
                    searchDO.PrReason,
                    searchDO.PrStatus);

                model.ProcessedPayments = this.webRepository.GetProcessedRequestsByUin(
                    this.CurrentUser.Uin,
                    searchDO.PrId,
                    Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                    Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                    searchDO.PrProvider,
                    searchDO.PrReason,
                    searchDO.PrStatus,
                    searchDO.PrSortBy,
                    searchDO.PrSortDesc,
                    searchDO.PrPage,
                    AppSettings.EPaymentsWeb_MaxSearchResultsPerPage);
            }
            else
            {
                model.PendingPayments = this.webRepository.GetPendingRequestsByAccessCode(this.CurrentUser.AccessCode);
                model.PendingPaymentsPagingOptions.TotalItemCount = model.PendingPayments.Count;

                model.ProcessedPayments = this.webRepository.GetProcessedRequestsByAccessCode(this.CurrentUser.AccessCode);
                model.ProcessedPaymentsPagingOptions.TotalItemCount = model.ProcessedPayments.Count;
            }

            model.DisabledItems = CheckPayment(model.PendingPayments);

            foreach (var m in model.PendingPayments)
            {
                var maxPayOrder = model.PendingPayments
                 .Where(p => p.PayOrder.HasValue && p.ObligationTypeId == m.ObligationTypeId && p.EserviceClientId == m.EserviceClientId).Max(x => x.PayOrder) ?? 0;
                m.MaxPayOrder = maxPayOrder;
            }

            AddExtendedConditionsToDisablePendingPayments(ref model);

            if (model.DisabledItems != null && model.DisabledItems.Any())
            {
                model.PendingPayments = DisableItems(model.DisabledItems, model.PendingPayments);
            }
            return View(model);
        }

        private void AddExtendedConditionsToDisablePendingPayments(ref ListVM model)
        {
            foreach (var processedPayment in model.ProcessedPayments)
            {
                try
                {
                    if (string.IsNullOrEmpty(processedPayment.AdditionalInformation))
                        continue;
                    var processedMdtInfo = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(processedPayment.AdditionalInformation));
                    if (processedPayment.PaymentRequestStatusId == PaymentRequestStatus.InProcess
                        && processedMdtInfo != null
                        && processedMdtInfo.TaxPeriodYear == DateTime.Now.Year
                        && processedMdtInfo.InstNo == 1)
                    {
                        model.DisabledItems.AddRange(
                            model.PendingPayments.Where(x =>
                            {
                                try
                                {
                                    if (string.IsNullOrEmpty(x.AdditionalInfo))
                                        return false;
                                    var pendingMdtInfo = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(x.AdditionalInfo));
                                    return pendingMdtInfo != null
                                        && x.PayOrder.HasValue
                                        && pendingMdtInfo.PartidaNo == processedMdtInfo.PartidaNo
                                        && pendingMdtInfo.InstNo > processedMdtInfo.InstNo
                                        && pendingMdtInfo.TaxPeriodYear == processedMdtInfo.TaxPeriodYear;
                                }
                                catch { return false; } // skip non json data in AdditionalInfo
                            }).Select(x => x.PaymentRequestIdentifier)
                        );
                    }
                }
                catch
                { continue; } // skip non json data in AdditionalInfo
            }
        }

        [HttpPost]
        [WebUserAuthorize]
        public virtual ActionResult ListSearch(ListSearchDO searchDO)
        {
            TempData[TempDataKeys.SearchPerformed] = true;

            searchDO.PrPage = 1;
            searchDO.Focus = Constants.ProcessedPaymentsFocusId;

            return RedirectToListAction(searchDO);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult ListSort(ListSearchDO searchDO)
        {
            return RedirectToListAction(searchDO);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult PaymentRequestsExportAsPdf(ListSearchDO searchDO)
        {
            ListVM model = new ListVM();

            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                model.PendingPayments = this.webRepository.GetAllPendingRequestsByUin(
                    this.CurrentUser.Uin,
                    searchDO.PSortBy,
                    searchDO.PSortDesc);

                model.ProcessedPayments = this.webRepository.GetAllProcessedRequestsByUin(
                    this.CurrentUser.Uin,
                    searchDO.PrId,
                    Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                    Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                    searchDO.PrProvider,
                    searchDO.PrReason,
                    searchDO.PrStatus,
                    searchDO.PrSortBy,
                    searchDO.PrSortDesc);
            }
            else
            {
                model.PendingPayments = this.webRepository.GetPendingRequestsByAccessCode(this.CurrentUser.AccessCode);

                model.ProcessedPayments = this.webRepository.GetProcessedRequestsByAccessCode(this.CurrentUser.AccessCode);
            }

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._UserPaymentRequestsPrintPdf, model);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "Задължения" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult PaymentRequestsExportAsExcel(ListSearchDO searchDO)
        {
            IList<PendingRequestVO> pendingPayments;
            IList<ProcessedRequestVO> processedPayments;

            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                pendingPayments = this.webRepository.GetAllPendingRequestsByUin(
                    this.CurrentUser.Uin,
                    searchDO.PSortBy,
                    searchDO.PSortDesc);

                processedPayments = this.webRepository.GetAllProcessedRequestsByUin(
                    this.CurrentUser.Uin,
                    searchDO.PrId,
                    Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                    Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                    searchDO.PrProvider,
                    searchDO.PrReason,
                    searchDO.PrStatus,
                    searchDO.PrSortBy,
                    searchDO.PrSortDesc);
            }
            else
            {
                pendingPayments = this.webRepository.GetPendingRequestsByAccessCode(this.CurrentUser.AccessCode);

                processedPayments = this.webRepository.GetProcessedRequestsByAccessCode(this.CurrentUser.AccessCode);
            }

            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Транзакции");

            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 6)).Merge();

            // starting from the second row
            int row = 1;
            int col = 1;
            IXLCell cell;

            // Print title
            // A1
            cell = worksheet.Cell(row, col);
            cell.Value = "Задължения за плащане";
            cell.Style.Font.Bold = true;

            // Move to the second row
            row = 2;

            #region Print Header
            // A2
            cell = worksheet.Cell(row, col);
            cell.Value = "Номер";
            cell.Style.Font.Bold = true;

            // B2
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Вид на задължението";
            cell.Style.Font.Bold = true;

            // C2
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Получател";
            cell.Style.Font.Bold = true;

            // D2
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Основание за плащане";
            cell.Style.Font.Bold = true;

            // E2
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дата и час";
            cell.Style.Font.Bold = true;

            // F2
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Валидно до";
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
            #endregion

            row += 1;
            col = 1;

            if (pendingPayments != null && pendingPayments.Count > 0)
            {
                for (int i = 0; i < pendingPayments.Count; i++)
                {
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(pendingPayments[i].PaymentRequestIdentifier);

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(pendingPayments[i].ObligationType);

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(pendingPayments[i].ServiceProviderName);

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(pendingPayments[i].PaymentReason);

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(Formatter.DateTimeToBgFormatWithoutSecondsNotLocalTime(pendingPayments[i].CreateDate));

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(Formatter.DateTimeToBgFormatWithoutSecondsNotLocalTime(pendingPayments[i].ExpirationDate));

                    if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(pendingPayments[i].PaymentAmount, pendingPayments[i].CreateDate));
                    }
                    else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
                    {
                        if (CurrencyHelper.IsEuroTimePeriod)
                        {
                            col++;
                            cell = worksheet.Cell(row, col);
                            cell.SetValue(CurrencyHelper.GetEuroValueFormated(pendingPayments[i].PaymentAmount, pendingPayments[i].CreateDate));

                            col++;
                            cell = worksheet.Cell(row, col);
                            cell.SetValue(CurrencyHelper.GetBgnValueFormated(pendingPayments[i].PaymentAmount, pendingPayments[i].CreateDate));
                        }
                        else
                        {
                            col++;
                            cell = worksheet.Cell(row, col);
                            cell.SetValue(CurrencyHelper.GetBgnValueFormated(pendingPayments[i].PaymentAmount, pendingPayments[i].CreateDate));

                            col++;
                            cell = worksheet.Cell(row, col);
                            cell.SetValue(CurrencyHelper.GetEuroValueFormated(pendingPayments[i].PaymentAmount, pendingPayments[i].CreateDate));
                        }
                    }
                    else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(pendingPayments[i].PaymentAmount, pendingPayments[i].CreateDate));
                    }

                    row++;
                    col = 1;
                }
            }
            else
            {
                worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, 6)).Merge();
                worksheet.Cell(row, 1).Value = "Няма намерени резултати";
            }

            // move one row down
            row++;
            col = 1;

            worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, 6)).Merge();

            cell = worksheet.Cell(row, col);
            cell.Value = "Последни движения";
            cell.Style.Font.Bold = true;

            // move one row down
            row++;

            #region Print headers

            cell = worksheet.Cell(row, col);
            cell.Value = "Номер";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Получател";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Вид на задължението";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Основание за плащане";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дата и час";
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
            cell.Value = "Статус";
            cell.Style.Font.Bold = true;

            #endregion

            // move to the start of the next row
            row++;
            col = 1;

            if (processedPayments != null && processedPayments.Count > 0)
            {
                for (int i = 0; i < processedPayments.Count; i++)
                {
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(processedPayments[i].PaymentRequestIdentifier);

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(processedPayments[i].ServiceProviderName);

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(processedPayments[i].ObligationType);

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(processedPayments[i].PaymentReason);

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(Formatter.DateTimeToBgFormatWithoutSecondsNotLocalTime(processedPayments[i].TransactionDate));

                    if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(processedPayments[i].PaymentAmountRequest, processedPayments[i].TransactionDate.GetValueOrDefault()));
                    }
                    else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
                    {
                        if (CurrencyHelper.IsEuroTimePeriod)
                        {
                            col++;
                            cell = worksheet.Cell(row, col);
                            cell.SetValue(CurrencyHelper.GetEuroValueFormated(processedPayments[i].PaymentAmountRequest, processedPayments[i].TransactionDate.GetValueOrDefault()));

                            col++;
                            cell = worksheet.Cell(row, col);
                            cell.SetValue(CurrencyHelper.GetBgnValueFormated(processedPayments[i].PaymentAmountRequest, processedPayments[i].TransactionDate.GetValueOrDefault()));
                        }
                        else
                        {
                            col++;
                            cell = worksheet.Cell(row, col);
                            cell.SetValue(CurrencyHelper.GetBgnValueFormated(processedPayments[i].PaymentAmountRequest, processedPayments[i].TransactionDate.GetValueOrDefault()));

                            col++;
                            cell = worksheet.Cell(row, col);
                            cell.SetValue(CurrencyHelper.GetEuroValueFormated(processedPayments[i].PaymentAmountRequest, processedPayments[i].TransactionDate.GetValueOrDefault()));
                        }
                    }
                    else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(processedPayments[i].PaymentAmountRequest, processedPayments[i].TransactionDate.GetValueOrDefault()));
                    }

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(Formatter.EnumToDescriptionString(processedPayments[i].PaymentRequestStatusId));

                    row++;
                    col = 1;
                }
            }
            else
            {
                worksheet.Range(worksheet.Cell(row, 1), worksheet.Cell(row, 6)).Merge();
                worksheet.Cell(row, 1).Value = "Няма намерени резултати";
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

            MemoryStream excelStream = new MemoryStream();
            workbook.SaveAs(excelStream);
            excelStream.Position = 0;

            return File(excelStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Задължения.xlsx");
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult Order(Guid id, bool showDetailsForm = false)
        {
            var paymentOrderVO = GetPaymentOrderVO(id);

            OrderVM model = new OrderVM(paymentOrderVO);

            //when (query param == true) then ShowDetailsForm is set to true, otherwise ShowDetailsForm is set in constructor
            if (showDetailsForm)
            {
                model.ShowDetailsForm = showDetailsForm;
            }

            return PartialView(MVC.Shared.Views._Order, model);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult Details(Guid id)
        {
            var paymentOrderVO = GetPaymentOrderVO(id);

            OrderVM model = new OrderVM(paymentOrderVO);

            return PartialView(model);
        }

        [HttpGet]
        [WebUserAuthorize]
        public virtual ActionResult CancelRequest(Guid id)
        {
            var paymentRequest = this.commonRepository.FindPaymentRequestByGuidAndUin(id, this.CurrentUser.Uin);

            if (paymentRequest.PaymentRequestStatusId == Model.Enums.PaymentRequestStatus.Pending)
            {
                paymentRequest.PaymentRequestStatusId = Model.Enums.PaymentRequestStatus.Canceled;
                paymentRequest.PaymentRequestStatusChangeTime = DateTime.Now;
                paymentRequest.ObligationStatusId = Model.Enums.ObligationStatusEnum.Canceled;

                this.unitOfWork.Save();

                User user = this.systemRepository.GetUserById(this.CurrentUser.UserId.Value);
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
            }

            return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult PrintOrder(Guid id)
        {
            var paymentOrderVO = GetPaymentOrderVO(id);

            OrderVM model = new OrderVM(paymentOrderVO);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._OrderPrint, model);

            return View(MVC.Shared.Views._PrintHtml, (object)htmlContent);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult PrintDetails(Guid id)
        {
            var paymentOrderVO = GetPaymentOrderVO(id);

            OrderVM model = new OrderVM(paymentOrderVO);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Payment.Views.DetailsPrint, model);

            return View(MVC.Shared.Views._PrintHtml, (object)htmlContent);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual FileResult DownloadPdfOrder(Guid id)
        {
            var paymentOrderVO = GetPaymentOrderVO(id);

            OrderVM model = new OrderVM(paymentOrderVO);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._OrderPrint, model);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "PaymentOrder" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual FileResult DownloadPdfDetails(Guid id)
        {
            var paymentOrderVO = GetPaymentOrderVO(id);

            OrderVM model = new OrderVM(paymentOrderVO);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Payment.Views.DetailsPrint, model);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "PaymentDetails" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult TaxAgreementDsk(Guid id)
        {
            PaymentRequest request;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                request = this.webRepository.GetPaymentRequestByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                request = this.webRepository.GetPaymentRequestByAccessCode(this.CurrentUser.AccessCode);
            }

            TaxAgreementDskVM model = new TaxAgreementDskVM();
            model.Gid = request.Gid;
            model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(request.PaymentAmount, request.CreateDate);
            model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(request.PaymentAmount, request.CreateDate);
            model.IsInternalPayment = true;

            return PartialView(MVC.Shared.Views._TaxAgreementDsk, model);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult TaxAgreementCvpos(Guid id)
        {
            PaymentRequest request;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                request = this.webRepository.GetPaymentRequestByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                request = this.webRepository.GetPaymentRequestByAccessCode(this.CurrentUser.AccessCode);
            }

            TaxAgreementCvposVM model = new TaxAgreementCvposVM();
            model.Gids = new Guid[] { request.Gid };
            model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(request.PaymentAmount, request.CreateDate);
            model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(request.PaymentAmount, request.CreateDate);
            model.IsInternalPayment = true;

            return PartialView(MVC.Shared.Views._TaxAgreementBoricaCvpos, model);
        }


        [HttpGet]
        public virtual ActionResult TaxAgreementCvposMultiple(Guid[] gids)
        {
            TaxAgreementCvposVM model = new TaxAgreementCvposVM();
            model.Gids = gids;

            List<PaymentRequest> paymentRequests = this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(pr => gids.Contains(pr.Gid)).ToList();

            foreach (var paymentRequest in paymentRequests)
            {
                model.PaymentAmountInBGN += CurrencyHelper.GetBgnValue(paymentRequest.PaymentAmount, paymentRequest.CreateDate);
                model.PaymentAmountInEuro += CurrencyHelper.GetEuroValue(paymentRequest.PaymentAmount, paymentRequest.CreateDate);
            }
            model.IsInternalPayment = true;

            return PartialView(MVC.Shared.Views._TaxAgreementBoricaCvpos, model);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult TaxAgreementBorica(Guid id)
        {
            PaymentRequest request;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                request = this.webRepository.GetPaymentRequestByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                request = this.webRepository.GetPaymentRequestByAccessCode(this.CurrentUser.AccessCode);
            }

            TaxAgreementBoricaVM model = new TaxAgreementBoricaVM();
            model.Gid = request.Gid;
            model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(request.PaymentAmount, request.CreateDate);
            model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(request.PaymentAmount, request.CreateDate);
            model.IsInternalPayment = true;

            return PartialView(MVC.Shared.Views._TaxAgreementBorica, model);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult TaxAgreementFiBank(Guid id)
        {
            PaymentRequest request;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                request = this.webRepository.GetPaymentRequestByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                request = this.webRepository.GetPaymentRequestByAccessCode(this.CurrentUser.AccessCode);
            }

            TaxAgreementFiBankVM model = new TaxAgreementFiBankVM();
            model.Gid = request.Gid;
            model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(request.PaymentAmount, request.CreateDate);
            model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(request.PaymentAmount, request.CreateDate);
            model.IsInternalPayment = true;

            return PartialView(MVC.Shared.Views._TaxAgreementFiBank, model);
        }

        [HttpGet]
        [WebUserAuthorize(AllowAuthorizationByAccessCode = true)]
        public virtual ActionResult TaxAgreementEpay(Guid id)
        {
            PaymentRequest request;
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                request = this.webRepository.GetPaymentRequestByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                request = this.webRepository.GetPaymentRequestByAccessCode(this.CurrentUser.AccessCode);
            }

            TaxAgreementEpayVM model = new TaxAgreementEpayVM();
            model.Gid = request.Gid;
            model.PaymentAmountInBGN = CurrencyHelper.GetBgnValue(request.PaymentAmount, request.CreateDate);
            model.PaymentAmountInEuro = CurrencyHelper.GetEuroValue(request.PaymentAmount, request.CreateDate);
            model.IsInternalPayment = true;

            return PartialView(MVC.Shared.Views._TaxAgreementEpay, model);
        }

        [HttpGet]
        [WebUserAuthorize]
        public virtual ActionResult Share(Guid id)
        {
            PaymentRequest request = this.webRepository.GetPaymentRequestByGidAndUin(id, this.CurrentUser.Uin);
            User user = this.systemRepository.GetUserById(this.CurrentUser.UserId.Value);

            if (String.IsNullOrWhiteSpace(request.PaymentRequestAccessCode))
            {
                request.PaymentRequestAccessCode = this.commonRepository.GeneratePaymentRequestAccessCode();
                request.ObligationStatusId = ObligationStatusEnum.Ordered;

                if (!string.IsNullOrWhiteSpace(user.Email))
                {
                    if (user.StatusObligationNotifications)
                    {
                        Email email = CreateObligationEmail(request, user);
                        this.systemRepository.AddEntity<Email>(email);

                        this.unitOfWork.Save();
                    }

                    if (user.AccessCodeNotifications)
                    {
                        AccessCodeApplicantActivatedContextDO contextObl = new AccessCodeApplicantActivatedContextDO(
                            request.PaymentRequestIdentifier,
                            request.ServiceProviderName,
                            request.PaymentReason,
                            request.PaymentAmount,
                            request.ObligationStatusId.GetDescription(),
                            request.PaymentRequestAccessCode,
                            request.CreateDate);

                        Email emailObl = new Email(contextObl, user.Email);
                        this.systemRepository.AddEntity<Email>(emailObl);
                    }
                }

                this.unitOfWork.Save();
            }

            if (request.ObligationStatusId == ObligationStatusEnum.Asked)
            {
                request.ObligationStatusId = ObligationStatusEnum.Ordered;

                if (!string.IsNullOrWhiteSpace(user.Email))
                {
                    if (user.StatusObligationNotifications)
                    {
                        Email email = CreateObligationEmail(request, user);
                        this.systemRepository.AddEntity<Email>(email);

                        this.unitOfWork.Save();
                    }
                }

                this.unitOfWork.Save();
            }


            ShareVM model = new ShareVM();

            model.Gid = id;
            model.AccessCode = request.PaymentRequestAccessCode;
            model.Email = user.Email;

            if (!String.IsNullOrWhiteSpace(request.PaymentRequestAccessCode))
            {
                model.Link = Formatter.UriCombine(AppSettings.EPaymentsCommon_WebAddress, MVC.Home.Name, MVC.Home.ActionNames.AccessByCode).ToString() + String.Format("?code={0}", request.PaymentRequestAccessCode);
            }

            return View(model);
        }

        [HttpPost]
        [WebUserAuthorize]
        public virtual ActionResult Share(ShareVM model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            PaymentRequest request = this.webRepository.GetPaymentRequestByGidAndUin(model.Gid, this.CurrentUser.Uin);

            SharePaymentContextDO contextDO = new SharePaymentContextDO(
                request.ApplicantName,
                request.PaymentRequestAccessCode,
                request.PaymentRequestIdentifier,
                request.ServiceProviderName,
                request.PaymentReason,
                request.PaymentAmount,
                request.CreateDate);

            Email email = new Email(contextDO, model.Email);
            this.systemRepository.AddEntity<Email>(email);
            this.unitOfWork.Save();

            TempData[TempDataKeys.SharePaymentEmailSend] = true;

            return RedirectToAction(MVC.Payment.ActionNames.Share, MVC.Payment.Name, new { id = request.Gid });
        }

        [HttpGet]
        [WebUserAuthorize]
        public virtual ActionResult GenerateAccessCode(Guid id)
        {
            PaymentRequest request = this.webRepository.GetPaymentRequestByGidAndUin(id, this.CurrentUser.Uin);
            request.PaymentRequestAccessCode = this.commonRepository.GeneratePaymentRequestAccessCode();

            User user = this.systemRepository.GetUserById(this.CurrentUser.UserId.Value);
            if (!String.IsNullOrWhiteSpace(user.Email) && user.AccessCodeNotifications)
            {
                AccessCodeActivatedContextDO contextDO = new AccessCodeActivatedContextDO(
                    request.ApplicantName,
                    request.PaymentRequestAccessCode,
                    request.PaymentRequestIdentifier,
                    request.ServiceProviderName,
                    request.PaymentReason,
                    request.PaymentAmount,
                    request.CreateDate);

                Email email = new Email(contextDO, user.Email);

                this.systemRepository.AddEntity<Email>(email);
            }

            this.unitOfWork.Save();

            return RedirectToAction(MVC.Payment.ActionNames.Share, MVC.Payment.Name, new { id = id });
        }


        [HttpGet]
        [WebUserAuthorize]
        public virtual JsonResult CheckItemRelatedPayments(string paymentRequestIdentifier)
        {
            var results = CheckRelatedPayments(paymentRequestIdentifier);
            return Json(results, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        [WebUserAuthorize]
        public virtual JsonResult CheckPaymentOnDeMark(string paymentRequestIdentifier)
        {
            //// returns all payments that should be de-marked.
            var res = CheckPayment(paymentRequestIdentifier);
            return Json(res, JsonRequestBehavior.AllowGet);
        }

        [NonAction]
        private ActionResult RedirectToListAction(ListSearchDO searchDO)
        {
            return RedirectToAction(MVC.Payment.ActionNames.List, MVC.Payment.Name,
                new
                {
                    @pPage = searchDO.PPage,
                    @pSortBy = searchDO.PSortBy,
                    @pSortDesc = searchDO.PSortDesc,

                    @prId = searchDO.PrId,
                    @prDateFrom = searchDO.PrDateFrom,
                    @prDateTo = searchDO.PrDateTo,
                    @prAmountFrom = searchDO.PrAmountFrom,
                    @prAmountTo = searchDO.PrAmountTo,
                    @prProvider = searchDO.PrProvider,
                    @prReason = searchDO.PrReason,
                    @prStatus = searchDO.PrStatus,

                    @prPage = searchDO.PrPage,
                    @prSortBy = searchDO.PrSortBy,
                    @prSortDesc = searchDO.PrSortDesc,

                    @focus = searchDO.Focus,
                });
        }

        [NonAction]
        private PaymentOrderVO GetPaymentOrderVO(Guid id)
        {
            if (!this.CurrentUser.IsAuthorizedByAccessCode)
            {
                return this.webRepository.GetPaymentOrderByGidAndUin(id, this.CurrentUser.Uin);
            }
            else
            {
                return this.webRepository.GetPaymentOrderByAccessCode(this.CurrentUser.AccessCode);
            }
        }

        [NonAction]
        private Email CreateObligationEmail(PaymentRequest request, User user)
        {
            StatusChangedObligationContextDO contextObl = new StatusChangedObligationContextDO(request.PaymentRequestIdentifier,
                request.ServiceProviderName,
                request.PaymentReason,
                request.PaymentAmount,
                request.ObligationStatusId.GetDescription(),
                request.CreateDate);

            Email email = new Email(contextObl, user.Email);

            return email;
        }

        private IList<PendingRequestVO> DisableItems(List<string> disabledItems, IList<PendingRequestVO> pendingItems)
        {
            pendingItems = pendingItems.Select(x =>
            {
                x.IsDisabled = disabledItems.Contains(x.PaymentRequestIdentifier);
                return x;
            })
                .ToList();
            return pendingItems;
        }
        private List<string> CheckPayment(IList<PendingRequestVO> pendingPayments)
        {
            List<MDT_ExtendedPaymentRequestDO> allData = new List<MDT_ExtendedPaymentRequestDO>();
            var paymentRequests = new List<PaymentRequest>();
            foreach (var x in pendingPayments)
            {
                var paymentRequest = this.webRepository.GetPaymentRequestByIdentifier(x.PaymentRequestIdentifier);
                paymentRequests.Add(paymentRequest);
            }
            var disabledItems = new List<string>();
            var eServiceClientIds = paymentRequests.Select(x => x.EserviceClientId).Distinct();
            foreach (var id in eServiceClientIds)
            {
                var paymentRequestList = this.webRepository
                                          .GetPendngAndInProcessPaymentRequestByAisClient(id, this.CurrentUser.Uin)
                                          .Where(pr => pr.PayOrder.HasValue)
                                          .Select(pr => new MDT_ExtendedPaymentRequestDO()
                                          {
                                              mDT_ExtendedInfoJson = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(pr.AdditionalInformation)),
                                              ObligationTypeId = pr.ObligationTypeId,
                                              PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                                              EserviceClientId = pr.EserviceClientId,
                                              PayOrder = pr.PayOrder,
                                          }).ToList();
                allData.AddRange(paymentRequestList);
            }

            foreach (var id in eServiceClientIds)
            {
                var obligationTypes = allData.Where(x => x.EserviceClientId == id).Select(x => x.ObligationTypeId).Distinct();

                foreach (var type in obligationTypes)
                {
                    var minPayOrder = allData.Where(p => p.PayOrder.HasValue && p.ObligationTypeId == type && p.EserviceClientId == id).Min(x => x.PayOrder) ?? 0;
                    var maxPayOrder = allData.Where(p => p.PayOrder.HasValue && p.ObligationTypeId == type && p.EserviceClientId == id).Max(x => x.PayOrder) ?? 0;

                    if (minPayOrder != maxPayOrder)
                    {
                        disabledItems.AddRange(allData.Where(p => p.PayOrder != minPayOrder && p.ObligationTypeId == type && p.EserviceClientId == id).Select(p => p.PaymentRequestIdentifier).ToList());
                    }
                    else
                    {
                        var allPartidas = allData.Where(p => p.PayOrder.HasValue && p.ObligationTypeId == type && p.EserviceClientId == id).Select(x => x.mDT_ExtendedInfoJson.PartidaNo).Distinct();
                        foreach (var partida in allPartidas)
                        {
                            var minInsta = allData
                                .Where(p => p.PayOrder.HasValue && p.ObligationTypeId == type && p.mDT_ExtendedInfoJson.PartidaNo == partida && p.EserviceClientId == id)
                                .Min(x => x.mDT_ExtendedInfoJson.InstNo);
                            var items = allData
                                .Where(p => p.PayOrder.HasValue &&
                                            p.EserviceClientId == id &&
                                            p.ObligationTypeId == type &&
                                            p.mDT_ExtendedInfoJson.PartidaNo == partida
                                            && p.mDT_ExtendedInfoJson.InstNo != minInsta)
                                .OrderBy(p => p.PayOrder)
                                .ThenBy(p => p.mDT_ExtendedInfoJson.InstNo)
                                .Select(p => p.PaymentRequestIdentifier)
                                .ToList();
                            disabledItems.AddRange(items);
                        }
                    }
                }
            }
            return disabledItems;
        }

        private List<string> CheckRelatedPayments(string id)
        {
            var paymentRequest = this.webRepository.GetPaymentRequestByIdentifier(id);
            var sourceMDTJson = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(paymentRequest.AdditionalInformation));

            var SessionCheckedPayments = this.webRepository.GetPendingRequestsByUin(
                    this.CurrentUser.Uin,
                    PendingPaymentColumn.PaymentId,
                    true);

            var paymentRequestList = this.webRepository
                                            .GetPendngPaymentRequestByAisClient(paymentRequest.EserviceClientId)
                                            .Where(p => p.ObligationTypeId == paymentRequest.ObligationTypeId)
                                            .Select(pr => new MDT_ExtendedPaymentRequestDO()
                                            {
                                                mDT_ExtendedInfoJson = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(pr.AdditionalInformation)),
                                                ObligationTypeId = pr.ObligationTypeId,
                                                PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                                                EserviceClientId = pr.EserviceClientId,
                                                PayOrder = pr.PayOrder,
                                            }).ToList();

            var res = new List<string>();
            var items = new List<string>();
            var maxOrder = paymentRequestList
                .Where(p => p.PayOrder.HasValue && p.ObligationTypeId == paymentRequest.ObligationTypeId && p.EserviceClientId == paymentRequest.EserviceClientId)
                .Max(x => x.PayOrder) ?? 0;

            if (paymentRequest.PayOrder != maxOrder)
            {
                var allPaymentsWithSamePayOrder = paymentRequestList
                    .Where(p => p.PayOrder.HasValue && p.ObligationTypeId == paymentRequest.ObligationTypeId && p.PayOrder == paymentRequest.PayOrder && p.EserviceClientId == paymentRequest.EserviceClientId)
                    .OrderBy(p => p.PayOrder)
                    .ToList();

                var paymentsWithHigherPayOrder = paymentRequestList
                    .Where(p => p.PayOrder.HasValue && p.ObligationTypeId == paymentRequest.ObligationTypeId && p.PayOrder > paymentRequest.PayOrder && p.EserviceClientId == paymentRequest.EserviceClientId)
                    .OrderBy(p => p.PayOrder)
                    .ToList();

                var internalMinOrder = paymentsWithHigherPayOrder
                    .Where(p => p.PayOrder.HasValue && p.ObligationTypeId == paymentRequest.ObligationTypeId && p.EserviceClientId == paymentRequest.EserviceClientId)
                    .Min(x => x.PayOrder) ?? 0;

                //// It is not latest payment.
                if (internalMinOrder != maxOrder)
                {
                    items = paymentsWithHigherPayOrder.Where(p => p.PayOrder == internalMinOrder && p.EserviceClientId == paymentRequest.EserviceClientId).Select(p => p.PaymentRequestIdentifier).ToList();
                    res.AddRange(items);
                }
                else
                {
                    //// It is latest payment
                    var partidaNos = paymentsWithHigherPayOrder.Select(x => x.mDT_ExtendedInfoJson.PartidaNo).Distinct();

                    foreach (var partidaNo in partidaNos)
                    {
                        var internalInstaMinOrder = paymentsWithHigherPayOrder.Where(p =>
                                                        p.PayOrder.HasValue &&
                                                        p.ObligationTypeId == paymentRequest.ObligationTypeId &&
                                                        p.PayOrder == maxOrder &&
                                                        p.EserviceClientId == paymentRequest.EserviceClientId &&
                                                        p.mDT_ExtendedInfoJson.PartidaNo == partidaNo
                                                    ).Min(x => x.mDT_ExtendedInfoJson.InstNo);
                        items = paymentsWithHigherPayOrder.Where(p =>
                                                        p.PayOrder == internalMinOrder &&
                                                        p.mDT_ExtendedInfoJson.InstNo == internalInstaMinOrder &&
                                                        p.EserviceClientId == paymentRequest.EserviceClientId &&
                                                        p.mDT_ExtendedInfoJson.PartidaNo == partidaNo
                                                    ).Select(p => p.PaymentRequestIdentifier).ToList();

                        res.AddRange(items);
                    }
                }
            }
            else
            {
                var latestPaymentData = paymentRequestList
                    .Where(p => p.ObligationTypeId == paymentRequest.ObligationTypeId && p.EserviceClientId == paymentRequest.EserviceClientId)
                    .Where(p => p.PayOrder.HasValue && p.PayOrder == paymentRequest.PayOrder)
                    .Where(p => p.mDT_ExtendedInfoJson.PartidaNo == sourceMDTJson.PartidaNo && p.mDT_ExtendedInfoJson.InstNo > sourceMDTJson.InstNo)
                    .OrderBy(p => p.PayOrder).ThenBy(p => p.mDT_ExtendedInfoJson.InstNo)
                    .Select(p => p.PaymentRequestIdentifier)
                    .Take(1)
                    .ToList();
                res = latestPaymentData;

            }

            return res;
        }

        private List<string> CheckPayment(string id)
        {

            var paymentRequest = this.webRepository.GetPaymentRequestByIdentifier(id);
            if (!paymentRequest.PayOrder.HasValue)
            {
                return null;
            }
            var sourceMDTJson = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(paymentRequest.AdditionalInformation));

            var paymentRequestList = this.webRepository
                                            .GetPendngPaymentRequestByAisClient(paymentRequest.EserviceClientId)
                                            .Select(pr => new MDT_ExtendedPaymentRequestDO()
                                            {
                                                mDT_ExtendedInfoJson = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(JsonHelper.RepairBrokenAdditionalInformationJsonString(pr.AdditionalInformation)),
                                                ObligationTypeId = pr.ObligationTypeId,
                                                PaymentRequestIdentifier = pr.PaymentRequestIdentifier,
                                                EserviceClientId = pr.EserviceClientId,
                                                PayOrder = pr.PayOrder,
                                            }).ToList();

            var maxOrder = paymentRequestList
                           .Where(p => p.PayOrder.HasValue && p.ObligationTypeId == paymentRequest.ObligationTypeId && p.EserviceClientId == paymentRequest.EserviceClientId).Max(x => x.PayOrder) ?? 0;

            var result = new List<string>();
            if (paymentRequest.PayOrder.Value != maxOrder)
            {
                result = paymentRequestList
                            .Where(p => p.ObligationTypeId == paymentRequest.ObligationTypeId && p.EserviceClientId == paymentRequest.EserviceClientId)
                            .Where(p => p.PayOrder.HasValue)
                            .Where(p => p.PayOrder > paymentRequest.PayOrder)
                            .Select(p => p.PaymentRequestIdentifier)
                            .ToList();
            }
            else
            {
                result = paymentRequestList
                     .Where(p => p.PayOrder.HasValue)
                     .Where(p => p.ObligationTypeId == paymentRequest.ObligationTypeId && p.PayOrder == paymentRequest.PayOrder && p.mDT_ExtendedInfoJson.PartidaNo == sourceMDTJson.PartidaNo && p.EserviceClientId == paymentRequest.EserviceClientId)
                      .Where(p => p.mDT_ExtendedInfoJson.InstNo > sourceMDTJson.InstNo)
                      .Select(p => p.PaymentRequestIdentifier)
                      .ToList();
            }
            result = result.Distinct().ToList();
            return result;
        }
    }
}