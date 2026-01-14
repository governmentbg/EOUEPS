using ClosedXML.Excel;
using EPayments.Admin.Auth;
using EPayments.Admin.Models.PaymentRequests;
using EPayments.Admin.Models.Shared;
using EPayments.Common;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Data.ViewObjects.Admin;
using EPayments.Model.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace EPayments.Admin.Controllers
{
    [AdminAuthorize(InternalAdminUserPermissionEnum.ViewReferences)]
    public partial class PaymentRequestController : BaseController
    {
        private readonly IPaymentRequestRepository PaymentRequestRepository;

        public PaymentRequestController(IPaymentRequestRepository paymentRequestRepository)
        {
            this.PaymentRequestRepository = paymentRequestRepository ?? throw new ArgumentNullException("paymentRequestRepository is null");
        }

        [HttpPost]
        public virtual ActionResult ListSearch(PaymentRequestSearchDO searchDO)
        {
            return RedirectToAction(MVC.PaymentRequest.ActionNames.List, MVC.PaymentRequest.Name,
                new
                {
                    @prId = searchDO.PrId,
                    @prRefenceNumber = searchDO.PrRefenceNumber,
                    @prDateFrom = searchDO.PrDateFrom,
                    @prDateTo = searchDO.PrDateTo,
                    @prAmountFrom = searchDO.PrAmountFrom,
                    @prAmountTo = searchDO.PrAmountTo,
                    @prProvider = searchDO.PrProvider,
                    @prReason = searchDO.PrReason,
                    @prApplicantName = searchDO.PrApplicantName,
                    @PrApplicantUin = searchDO.PrApplicantUin,
                    @prPaymentStatus = searchDO.PrPaymentStatus,
                    @prPaymentStatusChanged = searchDO.PrPaymentStatusChanged,
                    @prObligationStatus = searchDO.PrObligationStatus,

                    @prPage = searchDO.PrPage,
                    @prSortBy = searchDO.PrSortBy,
                    @prSortDesc = searchDO.PrSortDesc,

                    @focus = searchDO.Focus,
                });
        }

        [HttpGet]
        public virtual async Task<ActionResult> List(PaymentRequestSearchDO searchDO)
        {
            if (!searchDO.IsSearchForm && searchDO.PrPaymentStatusChanged.HasValue)
            {
                searchDO.PrPage = 1;

                await this.PaymentRequestRepository.ChangePaymentRequestsStatus(searchDO.PrId,
                    searchDO.PrRefenceNumber,
                    Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                    Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                    Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                    searchDO.PrProvider,
                    searchDO.PrReason,
                    searchDO.PrPaymentStatus,
                    searchDO.PrObligationStatus,
                    searchDO.PrApplicantName,
                    searchDO.PrApplicantUin,
                    searchDO.PrPaymentStatusChanged.Value);

                searchDO.IsSearchForm = true;
                await this.List();
            }
            PaymentRequestVM model = new PaymentRequestVM();

            model.RequestsPagingOptions = new PagingVM();

            model.SearchDO = searchDO;

            model.RequestsPagingOptions.CurrentPageIndex = searchDO.PrPage;
            model.RequestsPagingOptions.ControllerName = MVC.PaymentRequest.Name;
            model.RequestsPagingOptions.ActionName = MVC.PaymentRequest.ActionNames.List;
            model.RequestsPagingOptions.PageIndexParameterName = "prPage";
            model.RequestsPagingOptions.RouteValues = searchDO.ToRequestsRouteValues();

            model.RequestsPagingOptions.TotalItemCount = await this.PaymentRequestRepository.CountPaymentRequests(
                searchDO.PrId,
                searchDO.PrRefenceNumber,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                searchDO.PrProvider,
                searchDO.PrReason,
                searchDO.PrPaymentStatus,
                searchDO.PrObligationStatus,
                searchDO.PrApplicantName,
                searchDO.PrApplicantUin);

            model.Requests = await this.PaymentRequestRepository.GetPaymentRequests(
                searchDO.PrId,
                searchDO.PrRefenceNumber,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                searchDO.PrProvider,
                searchDO.PrReason,
                searchDO.PrPaymentStatus,
                searchDO.PrObligationStatus,
                searchDO.PrApplicantName,
                searchDO.PrApplicantUin,
                Enum.GetName(searchDO.PrSortBy.GetType(), searchDO.PrSortBy),
                searchDO.PrSortDesc,
                searchDO.PrPage,
                AppSettings.EPaymentsWeb_MaxSearchResultsPerPage);

            return View(model);
        }

        public virtual async Task<ActionResult> DownloadPdf(PaymentRequestSearchDO searchDO)
        {
            List<PaymentRequestVO> paymentRequests = await this.PaymentRequestRepository.GetAllRequests(searchDO.PrId,
                searchDO.PrRefenceNumber,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                searchDO.PrProvider,
                searchDO.PrReason,
                searchDO.PrPaymentStatus,
                searchDO.PrObligationStatus,
                searchDO.PrApplicantName,
                searchDO.PrApplicantUin,
                Enum.GetName(searchDO.PrSortBy.GetType(), searchDO.PrSortBy),
                searchDO.PrSortDesc,
                AppSettings.EPaymentsCommon_MaxPdfExportEntriesCount);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._PaymentRequestPrint,
                paymentRequests);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "Справка заявки за плащане" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        public virtual async Task<ActionResult> DownloadExcel(PaymentRequestSearchDO searchDO)
        {
            List<PaymentRequestVO> paymentRequests = await this.PaymentRequestRepository.GetAllRequests(searchDO.PrId,
                searchDO.PrRefenceNumber,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.PrDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.PrAmountTo),
                searchDO.PrProvider,
                searchDO.PrReason,
                searchDO.PrPaymentStatus,
                searchDO.PrObligationStatus,
                searchDO.PrApplicantName,
                searchDO.PrApplicantUin,
                Enum.GetName(searchDO.PrSortBy.GetType(), searchDO.PrSortBy),
                searchDO.PrSortDesc);

            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Справка заявки за плащане");

            int row = 1;
            int col = 1;
            IXLCell cell;

            // A1
            cell = worksheet.Cell(row, col);
            cell.Value = "Номер на задължението";
            cell.Style.Font.Bold = true;

            // B1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Референтен номер на задължението";
            cell.Style.Font.Bold = true;

            // C1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дата на създаване";
            cell.Style.Font.Bold = true;

            // D1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дата на изтичане";
            cell.Style.Font.Bold = true;

            // E1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Задължено лице";
            cell.Style.Font.Bold = true;

            // F1
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
            cell.Value = "Заявител";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Статус на плащането";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Статус на задължението";
            cell.Style.Font.Bold = true;

            // Move to the begining of row 2
            col = 1;
            row = 2;

            for (int i = 0; i < paymentRequests.Count; i++)
            {
                cell = worksheet.Cell(row, col);
                cell.SetValue(paymentRequests[i].PaymentRequestIdentifier);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(paymentRequests[i].PaymentReferenceNumber);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.DateToBgFormatWithoutYearSuffix(paymentRequests[i].CreateDate));

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.DateToBgFormatWithoutYearSuffix(paymentRequests[i].ExpirationDate));

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(paymentRequests[i].ApplicantName);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(paymentRequests[i].PaymentReason);

                if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetBgnValueFormated(paymentRequests[i].PaymentAmount, paymentRequests[i].CreateDate));
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
                {
                    if (CurrencyHelper.IsEuroTimePeriod)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(paymentRequests[i].PaymentAmount, paymentRequests[i].CreateDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(paymentRequests[i].PaymentAmount, paymentRequests[i].CreateDate));
                    }
                    else
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(paymentRequests[i].PaymentAmount, paymentRequests[i].CreateDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(paymentRequests[i].PaymentAmount, paymentRequests[i].CreateDate));
                    }
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(paymentRequests[i].PaymentAmount, paymentRequests[i].CreateDate));
                }

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(paymentRequests[i].ServiceProviderName);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.EnumToDescriptionString(paymentRequests[i].PaymentRequestStatusId));

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(paymentRequests[i].ObligationStatusId != null ? Formatter.EnumToDescriptionString(paymentRequests[i].ObligationStatusId) : "Няма стойност");

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

            MemoryStream excelStream = new MemoryStream();
            workbook.SaveAs(excelStream);
            excelStream.Position = 0;

            return File(excelStream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Справка заявки за плащане.xlsx");
        }
    }
}