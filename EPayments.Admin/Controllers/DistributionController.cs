using ClosedXML.Excel;
using EPayments.Admin.Auth;
using EPayments.Admin.Models.Distributions;
using EPayments.Admin.Models.Shared;
using EPayments.Common;
using EPayments.Common.Helpers;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Data.ViewObjects.Admin;
using EPayments.Distributions.Interfaces;
using EPayments.Distributions.Models.BNB;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using log4net;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using System.Xml.Linq;

namespace EPayments.Admin.Controllers
{
    [AdminAuthorize(InternalAdminUserPermissionEnum.DistributionReferences)]
    public partial class DistributionController : BaseController
    {
        private readonly IDistributionRepository DistributionRepository;
        private readonly IDistributionFactory DistributionFactory;
        private static readonly ILog Logger = LogManager.GetLogger(nameof(DistributionController));

        public DistributionController(IDistributionRepository distributionRepository,
            IDistributionFactory distributionFactory)
        {
            this.DistributionRepository = distributionRepository ?? throw new ArgumentNullException("distributionRepository is null.");
            this.DistributionFactory = distributionFactory ?? throw new ArgumentNullException("distributionFactory is null");
        }

        public virtual async Task<ActionResult> Distributions(DistributionRevenueSearchDO searchDO)
        {
            if (searchDO == null)
            {
                searchDO = new DistributionRevenueSearchDO();
            }
            DistributionRevenueVM model = new DistributionRevenueVM();
            try
            {
                model.RequestsPagingOptions = new PagingVM();
                model.RequestsPagingOptions.ActionName = MVC.Distribution.ActionNames.Distributions;
                model.RequestsPagingOptions.ControllerName = MVC.Distribution.Name;
                model.RequestsPagingOptions.CurrentPageIndex = searchDO.CurrentPage;
                model.RequestsPagingOptions.PageIndexParameterName = searchDO.PageIndexParameterName;
                model.RequestsPagingOptions.RouteValues = searchDO.ToDistributionRouteValues(searchDO.SortBy);
                model.RequestsPagingOptions.TotalItemCount = await this.CountDistributions(searchDO);
                model.DistributionRevenues = await this.GetDistributions(searchDO, model.RequestsPagingOptions.PageSize);
                model.DistribtutionTypes = await this.GetDistribtutionTypes();
                model.SearchDO = searchDO;
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }

            return View(model);
        }

        public virtual async Task<ActionResult> Payments(PaymentSearchDO searchDO)
        {
            DistributionRevenueVO distributionRevenue = await this.DistributionRepository.GetDistributionById(searchDO.Id);

            if (distributionRevenue == null)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Разпределението не е намерено.";

                return this.RedirectToAction(MVC.Distribution.ActionNames.Distributions, MVC.Distribution.Name);
            }

            PaymentVM model = new PaymentVM();

            model.SearchDO = searchDO;
            model.DistributionRevenue = distributionRevenue;
            model.Payments = await this.DistributionRepository.GetDistributionPaymentRequests(searchDO.Id,
            Enum.GetName(searchDO.SortBy.GetType(), searchDO.SortBy),
            searchDO.SortDesc);
            model.DistribtutionTypes = await this.GetDistribtutionTypes();

            return View(model);
        }

        public virtual async Task<ActionResult> DownloadPdf(PaymentSearchDO searchDO)
        {
            DistributionRevenueVO distributionRevenue = await this.DistributionRepository.GetDistributionById(searchDO.Id);

            if (distributionRevenue == null)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Разпределението не е намерено.";

                return this.RedirectToAction(MVC.Distribution.ActionNames.Distributions, MVC.Distribution.Name);
            }

            PaymentVM model = new PaymentVM();

            model.DistributionRevenue = distributionRevenue;

            model.Payments = await this.DistributionRepository.GetAllDistributionPaymentRequests(searchDO.Id,
            Enum.GetName(searchDO.SortBy.GetType(), searchDO.SortBy),
            searchDO.SortDesc, AppSettings.EPaymentsCommon_MaxPdfExportEntriesCount);
            model.DistribtutionTypes = await this.GetDistribtutionTypes();

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._DistributionRevenuePrintPdf, model);

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "Справка разпределение" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        public virtual async Task<ActionResult> DownloadExcel(PaymentSearchDO searchDO)
        {
            DistributionRevenueVO distributionRevenue = await this.DistributionRepository.GetDistributionById(searchDO.Id);

            if (distributionRevenue == null)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Разпределението не е намерено.";

                return this.RedirectToAction(MVC.Distribution.ActionNames.Distributions, MVC.Distribution.Name);
            }

            PaymentVM model = new PaymentVM();

            model.DistributionRevenue = distributionRevenue;

            model.Payments = await this.DistributionRepository.GetAllDistributionPaymentRequests(searchDO.Id,
            Enum.GetName(searchDO.SortBy.GetType(), searchDO.SortBy),
            searchDO.SortDesc);

            model.DistribtutionTypes = await this.GetDistribtutionTypes();

            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Справка заявки за плащане");

            int row = 1;
            int col = 1;
            IXLCell cell;

            // A1
            cell = worksheet.Cell(row, col);
            cell.Value = "Справката е генерирана на";
            cell.Style.Font.Bold = true;

            // B1
            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Справката е разпределена на";
            cell.Style.Font.Bold = true;

            if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
            {
                // C1
                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Обща сума на разпределението, BGN";
                cell.Style.Font.Bold = true;
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
            {
                if (CurrencyHelper.IsEuroTimePeriod)
                {
                    // C1
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Обща сума на разпределението, EUR";
                    cell.Style.Font.Bold = true;

                    // D1
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Обща сума на разпределението, BGN";
                    cell.Style.Font.Bold = true;
                }
                else
                {
                    // C1
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Обща сума на разпределението, BGN";
                    cell.Style.Font.Bold = true;

                    // D1
                    col++;
                    cell = worksheet.Cell(row, col);
                    worksheet.Cell(row, col).Value = "Обща сума на разпределението, EUR";
                    cell.Style.Font.Bold = true;
                }
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
            {
                // D1
                col++;
                cell = worksheet.Cell(row, col);
                worksheet.Cell(row, col).Value = "Обща сума на разпределението, EUR";
                cell.Style.Font.Bold = true;
            }

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дали е разпределенa";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Изпратена към Борика";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Вид на разпределението";
            cell.Style.Font.Bold = true;

            // Move to the begining of the second row
            row = 2;
            col = 1;

            // A2
            cell = worksheet.Cell(row, col);
            cell.SetValue(Formatter.DateTimeToBgFormatWithoutSeconds(model.DistributionRevenue.CreatedAt));

            // B2
            col++;
            cell = worksheet.Cell(row, col);
            cell.SetValue(Formatter.DateTimeToBgFormatWithoutSeconds(model.DistributionRevenue.DistributedDate) ?? "не е разпределено");

            if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
            {
                // C2
                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(CurrencyHelper.GetBgnValueFormated(model.DistributionRevenue.TotalSum, model.DistributionRevenue.DistributedDate.GetValueOrDefault()));
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
            {
                if (CurrencyHelper.IsEuroTimePeriod)
                {
                    // C2
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(model.DistributionRevenue.TotalSum, model.DistributionRevenue.DistributedDate.GetValueOrDefault()));

                    // D2
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue<string>(CurrencyHelper.GetBgnValueFormated(model.DistributionRevenue.TotalSum, model.DistributionRevenue.DistributedDate.GetValueOrDefault()));
                }
                else
                {
                    // C2
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetBgnValueFormated(model.DistributionRevenue.TotalSum, model.DistributionRevenue.DistributedDate.GetValueOrDefault()));

                    // D2
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(model.DistributionRevenue.TotalSum, model.DistributionRevenue.DistributedDate.GetValueOrDefault()));
                }
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
            {
                // C2
                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(CurrencyHelper.GetEuroValueFormated(model.DistributionRevenue.TotalSum, model.DistributionRevenue.DistributedDate.GetValueOrDefault()));
            }

            // Дали е разпределенa
            col++;
            cell = worksheet.Cell(row, col);
            cell.SetValue(model.DistributionRevenue.IsDistributed ? "Да" : "Не");

            // Изпратена към Борика
            col++;
            cell = worksheet.Cell(row, col);
            cell.SetValue(model.DistributionRevenue.IsFileGenerated ? "Да" : "Не");

            // Вид на разпределението
            col++;
            cell = worksheet.Cell(row, col);
            cell.SetValue(model.DistribtutionTypes.FirstOrDefault(dt => dt.DistributionTypeId == model.DistributionRevenue.DistributionType)?.Name ?? string.Empty);

            // Move to the begining of row 4
            col = 1;
            row = 4;

            cell = worksheet.Cell(row, col);
            cell.Value = "Задължения в разпределението";
            cell.Style.Font.Bold = true;

            worksheet.Range(worksheet.Cell(4, 1), worksheet.Cell(4, 8)).Merge();

            // Move to the begining of row 5
            col = 1;
            row = 5;

            cell = worksheet.Cell(row, col);
            cell.Value = "Номер на задължението";
            cell.Style.Font.Bold = true;

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
            cell.Value = "АИС разпоредител";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Задължено лице";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Статус на плащането";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Статус на задължението";
            cell.Style.Font.Bold = true;

            // Move to the begining of row 5
            col = 1;
            row = 6;

            for (int i = 0; i < model.Payments.Count; i++)
            {
                cell = worksheet.Cell(row, col);
                cell.SetValue(model.Payments[i].PaymentRequestIdentifier);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(model.Payments[i].PaymentReason);

                if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetBgnValueFormated(model.Payments[i].PaymentAmount, model.Payments[i].CreateDate));
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
                {
                    if (CurrencyHelper.IsEuroTimePeriod)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(model.Payments[i].PaymentAmount, model.Payments[i].CreateDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(model.Payments[i].PaymentAmount, model.Payments[i].CreateDate));
                    }
                    else
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(model.Payments[i].PaymentAmount, model.Payments[i].CreateDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(model.Payments[i].PaymentAmount, model.Payments[i].CreateDate));
                    }
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(model.Payments[i].PaymentAmount, model.Payments[i].CreateDate));
                }

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(model.Payments[i].EServiceClientName);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(model.Payments[i].TargetEServiceClientName);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(model.Payments[i].ApplicantName);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.EnumToDescriptionString(model.Payments[i].PaymentRequestStatus));

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(model.Payments[i].ObligationStatus != null ? Formatter.EnumToDescriptionString(model.Payments[i].ObligationStatus) : "Няма стойност");

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

            MemoryStream excelStream = new MemoryStream();

            workbook.SaveAs(excelStream);
            excelStream.Position = 0;

            return File(excelStream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Справка разпределение.xlsx");
        }

        public virtual async Task<ActionResult> Distribute(int id)
        {
            DistributionRevenue distributionRevenue = await DistributionRepository.GetDistribution(id);

            if (distributionRevenue == null)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Разпределението не е намерено.";

                return this.RedirectToAction(MVC.Distribution.ActionNames.Distributions, MVC.Distribution.Name);
            }

            if (distributionRevenue.IsDistributed == true)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Разпределението вече е било разпределено.";

                return this.RedirectToAction(MVC.Distribution.ActionNames.Distributions, MVC.Distribution.Name);
            }

            List<PaymentRequest> paymentRequests = await DistributionRepository.GetDistributionPaymentRequests(id);

            distributionRevenue.IsDistributed = true;

            paymentRequests.ForEach(pr => pr.ObligationStatusId = ObligationStatusEnum.CheckedAccount);

            this.DistributionRepository.Save();

            TempData[Common.TempDataKeys.Message] = "Разпределението e маркирано като разпределено.";

            return this.RedirectToAction(MVC.Distribution.ActionNames.Payments, MVC.Distribution.Name, new { id = id });
        }

        public virtual async Task<ActionResult> GetFile(int id)
        {
            DistributionRevenue distributionRevenue = await DistributionRepository.GetDistribution(id);

            if (distributionRevenue == null)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Разпределението не е намерено.";

                return this.RedirectToAction(MVC.Distribution.ActionNames.Distributions, MVC.Distribution.Name);
            }

            string bulstat = AppSettings.EPaymentsJobHost_DistributionBulstat;
            string bicCode = AppSettings.EPaymentsJobHost_DistributionBICCode;
            string eGov = AppSettings.EPaymentsJobHost_DistributionSenderName;
            string iban = AppSettings.EPaymentsJobHost_DistributionIban;
            string xsdName = AppSettings.EPaymentsJobHost_XsdFileName;
            string vpn = AppSettings.EPaymentsJobHost_Vpn;
            string vd = AppSettings.EPaymentsJobHost_Vd;
            string firstDescription = AppSettings.EPaymentsJobHost_FirstDescription;
            string secondDescription = AppSettings.EPaymentsJobHost_SecondDescription;
            string xsdDirectoryPath = AppSettings.EPaymentsJobHost_SchemasDirectory;
            var obligationTypeList = await DistributionRepository.GetAllObligationTypes();

            distributionRevenue.DistributionRevenuePayments = await DistributionRepository
                .GetDistributionRevenuePayments(distributionRevenue.DistributionRevenueId);

            BnbFile bnbFile = this.DistributionFactory.BnbModelCreator()
                           .Create(distributionRevenue, bulstat, eGov, iban, bicCode, vpn, vd, firstDescription, secondDescription, obligationTypeList);

            IBnbXmlDocumentCreator bnbXmlDocumentCreator = this.DistributionFactory.BnbXmlDocumentCreator();

            XDocument document = bnbXmlDocumentCreator.CreateDocument(bnbFile);

            List<string> errors = bnbXmlDocumentCreator.ValidateDocument(document, xsdDirectoryPath, xsdName);

            if (errors.Count > 0)
            {
                errors.ForEach(e =>
                {
                    if (!string.IsNullOrWhiteSpace(e) && !distributionRevenue.DistributionErrors.Any(de => string.Equals(de.Error, e, StringComparison.OrdinalIgnoreCase)))
                    {
                        distributionRevenue.DistributionErrors.Add(new DistributionError()
                        {
                            CreatedAt = DateTime.Now.ToUniversalTime(),
                            Error = e.Length > 500 ? e.Substring(0, 500) : e
                        });
                    }
                });

                DistributionRepository.Save();
            }

            MemoryStream xmlStream = new MemoryStream();

            document.Save(xmlStream);

            xmlStream.Position = 0;

            return File(xmlStream, "application/xml", string.Format("{0}-{1}.xml", distributionRevenue.DistributionRevenueId, Formatter.DateToBgFormatWithoutYearSuffix(DateTime.Today)));
        }

        private async Task<List<DistributionRevenueVO>> GetDistributions(DistributionRevenueSearchDO searchDO, int pageLength)
        {
            return await this.DistributionRepository.GetDistributionRevenues(
                searchDO.CurrentPage,
                pageLength,
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.StartDate)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.EndDate)),
                searchDO.DistributionType,
                searchDO.DistributionRevenueId,
                Enum.GetName(searchDO.SortBy.GetType(), searchDO.SortBy),
                searchDO.SortDesc);
        }

        private async Task<int> CountDistributions(DistributionRevenueSearchDO searchDO)
        {
            return await this.DistributionRepository.CountDistributionRevenues(
                Parser.GetDateFirstMinute(Parser.BgFormatDateStringToDateTime(searchDO.StartDate)),
                Parser.GetDateLastMinute(Parser.BgFormatDateStringToDateTime(searchDO.EndDate)),
                searchDO.DistributionType);
        }

        private async Task<List<DistribtutionTypeVO>> GetDistribtutionTypes()
        {
            return await this.DistributionRepository.GetDistributionTypes();
        }
    }
}