using ClosedXML.Excel;
using EPayments.Admin.Auth;
using EPayments.Admin.Models.Shared;
using EPayments.Admin.Models.UndistributedPayments;
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
    public partial class UndistributedPaymentsController : BaseController
    {
        private readonly IEquationControlsRepository EquationControlsRepository;

        public UndistributedPaymentsController(IEquationControlsRepository equationControlsRepository)
        {
            this.EquationControlsRepository = equationControlsRepository ?? throw new ArgumentNullException("equationControlsRepository is null");
        }

        [HttpGet]
        public virtual async Task<ActionResult> List(UndistributedPaymentSearchDO searchDO)
        {
            UndistributedPaymentVM model = new UndistributedPaymentVM();

            model.RequestsPagingOptions = new PagingVM();
            model.SearchDO = searchDO;

            model.RequestsPagingOptions.CurrentPageIndex = searchDO.UpPage;
            model.RequestsPagingOptions.ControllerName = MVC.UndistributedPayments.Name;
            model.RequestsPagingOptions.ActionName = MVC.UndistributedPayments.ActionNames.List;
            model.RequestsPagingOptions.PageIndexParameterName = "upPage";
            model.RequestsPagingOptions.RouteValues = searchDO.ToRequestsRouteValues();

            model.RequestsPagingOptions.TotalItemCount = await this.EquationControlsRepository.CountUndistributetPayments(
                searchDO.UpId,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.UpDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.UpDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.UpAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.UpAmountTo),
                searchDO.UpProvider,
                searchDO.UpReason,
                searchDO.UpObligationStatus);

            model.Requests = await this.EquationControlsRepository.GetUndistributetPayments(
                searchDO.UpId,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.UpDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.UpDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.UpAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.UpAmountTo),
                searchDO.UpProvider,
                searchDO.UpReason,
                searchDO.UpObligationStatus,
                Enum.GetName(searchDO.UpSortBy.GetType(), searchDO.UpSortBy),
                searchDO.UpSortDesc,
                searchDO.UpPage,
                AppSettings.EPaymentsWeb_MaxSearchResultsPerPage);

            return View(model);
        }

        public virtual async Task<ActionResult> DownloadPdf(UndistributedPaymentSearchDO searchDO)
        {
            List<PaymentRequestVO> model = await this.EquationControlsRepository.GetAllUndistributetPayments(
                searchDO.UpId,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.UpDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.UpDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.UpAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.UpAmountTo),
                searchDO.UpProvider,
                searchDO.UpReason,
                searchDO.UpObligationStatus,
                Enum.GetName(searchDO.UpSortBy.GetType(), searchDO.UpSortBy),
                searchDO.UpSortDesc,
                AppSettings.EPaymentsCommon_MaxPdfExportEntriesCount);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._UndistributedPaymentsPrintPdf, model);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "Контрола за равнение по неразпределени задължения" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        public virtual async Task<ActionResult> DownloadExcel(UndistributedPaymentSearchDO searchDO)
        {
            List<PaymentRequestVO> model = await this.EquationControlsRepository.GetAllUndistributetPayments(
                searchDO.UpId,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.UpDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.UpDateTo)),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.UpAmountFrom),
                Parser.TwoDecimalPlacesFormatStringToDecimal(searchDO.UpAmountTo),
                searchDO.UpProvider,
                searchDO.UpReason,
                searchDO.UpObligationStatus,
                Enum.GetName(searchDO.UpSortBy.GetType(), searchDO.UpSortBy),
                searchDO.UpSortDesc);

            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Неразпределени задължения");

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
            cell.Value = "Дата на създаване";
            cell.Style.Font.Bold = true;

            if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
            {
                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Сума, BGN";
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
            {
                if (CurrencyHelper.IsEuroTimePeriod)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Сума, EUR";

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Сума, BGN";
                }
                else
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Сума, BGN";

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Сума, EUR";
                }
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
            {
                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Сума, EUR";
            }

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Заявител";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Статус на задължението";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Причина за неразпределение";
            cell.Style.Font.Bold = true;

            // Move to the begining of the second row
            col = 1;
            row = 2;

            for (int i = 0; i < model.Count; i++)
            {
                cell = worksheet.Cell(row, col);
                cell.SetValue(model[i].PaymentRequestIdentifier);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.DateToBgFormatWithoutYearSuffix(model[i].CreateDate));

                if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetBgnValueFormated(model[i].PaymentAmount, model[i].CreateDate));
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
                {
                    if (CurrencyHelper.IsEuroTimePeriod)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(model[i].PaymentAmount, model[i].CreateDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(model[i].PaymentAmount, model[i].CreateDate));

                    }
                    else
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(model[i].PaymentAmount, model[i].CreateDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(model[i].PaymentAmount, model[i].CreateDate));
                    }
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(model[i].PaymentAmount, model[i].CreateDate));
                }

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(model[i].ServiceProviderName);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(model[i].ObligationStatusId != null ? Formatter.EnumToDescriptionString(model[i].ObligationStatusId) : "Няма стойност");

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(model[i].PaymentReason);

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

            MemoryStream excelStream = new MemoryStream();
            workbook.SaveAs(excelStream);
            excelStream.Position = 0;

            return File(excelStream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Контрола за равнение по неразпределени задължения.xlsx");
        }
    }
}