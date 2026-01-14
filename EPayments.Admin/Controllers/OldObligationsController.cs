using EPayments.Admin.Auth;
using EPayments.Admin.Models.Shared;
using EPayments.Admin.Models.OldObligations;
using EPayments.Common;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Model.Enums;
using System;
using System.Threading.Tasks;
using System.Web.Mvc;
using EPayments.Data.ViewObjects.Admin;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.IO;
using System.Linq;

namespace EPayments.Admin.Controllers
{
    [AdminAuthorize(InternalAdminUserPermissionEnum.ViewReferences)]
    public partial class OldObligationsController : BaseController
    {
        private readonly IEquationControlsRepository EquationControlsRepository;

        public OldObligationsController(IEquationControlsRepository equationControlsRepository)
        {
            this.EquationControlsRepository = equationControlsRepository ?? throw new ArgumentNullException("equationControlsRepository is null");
        }

        [HttpGet]
        public virtual async Task<ActionResult> List(OldObligationsSearchDO searchDO)
        {
            OldObligationsVM model = new OldObligationsVM();

            model.RequestsPagingOptions = new PagingVM();
            model.SearchDO = searchDO;
            if (model.SearchDO.OoObligationStatus == null)
                model.SearchDO.OoObligationStatus = ObligationStatusEnum.IrrevocableOrder;
            model.RequestsPagingOptions.CurrentPageIndex = searchDO.OoPage;
            model.RequestsPagingOptions.ControllerName = MVC.OldObligations.Name;
            model.RequestsPagingOptions.ActionName = MVC.OldObligations.ActionNames.List;
            model.RequestsPagingOptions.PageIndexParameterName = "ooPage";
            model.RequestsPagingOptions.RouteValues = searchDO.ToRequestsRouteValues();

            model.RequestsPagingOptions.TotalItemCount = await this.EquationControlsRepository.CountOldPayments(
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.OoDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.OoDateTo)),
                searchDO.OoObligationStatus);

            model.Requests = await this.EquationControlsRepository.GetOldPayments(
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.OoDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.OoDateTo)),
                searchDO.OoObligationStatus,
                Enum.GetName(searchDO.OoSortBy.GetType(), searchDO.OoSortBy),
                searchDO.OoSortDesc,
                searchDO.OoPage,
                AppSettings.EPaymentsWeb_MaxSearchResultsPerPage);

            return View(model);
        }

        public virtual async Task<ActionResult> DownloadPdf(OldObligationsSearchDO searchDO)
        {
            List<PaymentRequestVO> model = await this.EquationControlsRepository.GetAllOldPayments(
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.OoDateFrom)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.OoDateTo)),
                searchDO.OoObligationStatus,
                Enum.GetName(searchDO.OoSortBy.GetType(), searchDO.OoSortBy),
                searchDO.OoSortDesc,
                AppSettings.EPaymentsCommon_MaxPdfExportEntriesCount);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._OldObligationsPaymentPrintPdf, model);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "Контрол за равнение по стари задължения" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        public virtual async Task<ActionResult> DownloadExcel(OldObligationsSearchDO searchDO)
        {
            List<PaymentRequestVO> model = await this.EquationControlsRepository.GetAllOldPayments(
                 Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.OoDateFrom)),
                 Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.OoDateTo)),
                 searchDO.OoObligationStatus,
                 Enum.GetName(searchDO.OoSortBy.GetType(), searchDO.OoSortBy),
                 searchDO.OoSortDesc);

            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Контрол стари задължения");

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
            cell.Value = "Статус на задължението";
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

            MemoryStream excelStream = new MemoryStream();
            workbook.SaveAs(excelStream);
            excelStream.Position = 0;

            return File(excelStream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Контрол за равнение по стари задължения.xlsx");
        }
    }
}