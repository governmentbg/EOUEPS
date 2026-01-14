using ClosedXML.Excel;
using DocumentFormat.OpenXml.EMMA;
using EPayments.Admin.Auth;
using EPayments.Admin.Models.Shared;
using EPayments.Admin.Models.Transactions;
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
    public partial class TransactionController : BaseController
    {
        private readonly IEquationControlsRepository EquationControlsRepository;

        public TransactionController(IEquationControlsRepository equationControlsRepository)
        {
            this.EquationControlsRepository = equationControlsRepository ?? throw new ArgumentNullException("equationControlsRepository is null");
        }

        public virtual async Task<ActionResult> List(TransactionSearchDO searchDO)
        {
            TransactionVM model = new TransactionVM();

            model.RequestsPagingOptions = new PagingVM();
            model.SearchDO = searchDO;

            model.RequestsPagingOptions.CurrentPageIndex = searchDO.TtPage;
            model.RequestsPagingOptions.ControllerName = MVC.Transaction.Name;
            model.RequestsPagingOptions.ActionName = MVC.Transaction.ActionNames.List;
            model.RequestsPagingOptions.PageIndexParameterName = "TtPage";
            model.RequestsPagingOptions.RouteValues = searchDO.ToRequestsRouteValues();

            DateTime? dateFrom = Parser.BgFormatDateStringToDateTime(searchDO.TtDateFrom);
            DateTime? dateTo = Parser.BgFormatDateStringToDateTime(searchDO.TtDateTo);

            int? transactionStatus;

            if (searchDO.TtTransactionStatus.HasValue)
            {
                transactionStatus = (int)searchDO.TtTransactionStatus.Value;
            }
            else
            {
                transactionStatus = null;
            }

            var transactions = await this.EquationControlsRepository.GetBoricaTransactions(
                dateFrom,
                dateTo,
                transactionStatus,
                Enum.GetName(searchDO.TtSortBy.GetType(), searchDO.TtSortBy),
                searchDO.TtSortDesc,
                searchDO.TtPage,
                AppSettings.EPaymentsWeb_MaxSearchResultsPerPage);

            model.Transactions = transactions;

            var totals = await EquationControlsRepository.CountBoricaTransactions(dateFrom, dateTo, transactionStatus);
            model.RequestsPagingOptions.TotalItemCount = totals.TotalPages;
            model.CalculateTotalAmountInEuro = totals.TotalAmountInEuro;
            model.CalculateTotalFeeInEuro = totals.TotalFeeInEuro;
            model.CalculateTotalCommissionInEuro = totals.CommissionInEuro;

            return View(model);
        }

        public virtual async Task<ActionResult> Payments(int transactionId)
        {
            if (transactionId <= 0)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Транзакцията не е намерена.";

                return this.RedirectToAction(MVC.Transaction.ActionNames.List, MVC.Transaction.Name);
            }

            BoricaTransactionVO boricaTransaction = await this.EquationControlsRepository.GetBoricaTransaction(transactionId);

            if (boricaTransaction == null)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Транзакцията не е намерена.";

                return this.RedirectToAction(MVC.Transaction.ActionNames.List, MVC.Transaction.Name);
            }

            List<PaymentRequestVO> paymentRequests = await this.
                EquationControlsRepository.GetTransactionPaymentRequests(transactionId);

            return View(new TransactionWithPaymentsVM()
            {
                BoricaTransaction = boricaTransaction,
                Payments = paymentRequests
            });
        }

        public virtual async Task<ActionResult> DownloadExcel(TransactionSearchDO searchDO)
        {
            DateTime? dateFrom = Parser.BgFormatDateStringToDateTime(searchDO.TtDateFrom);
            DateTime? dateTo = Parser.BgFormatDateStringToDateTime(searchDO.TtDateTo);

            DateTime defaultDate = new DateTime();

            int? transactionStatus;

            if (searchDO.TtTransactionStatus.HasValue)
            {
                transactionStatus = (int)searchDO.TtTransactionStatus.Value;
            }
            else
            {
                transactionStatus = null;
            }

            List<BoricaTransactionVO> transactions = await this.EquationControlsRepository.GetBoricaTransactions(
                dateFrom,
                dateTo,
                transactionStatus,
                Enum.GetName(searchDO.TtSortBy.GetType(), searchDO.TtSortBy),
                searchDO.TtSortDesc);

            decimal calculateTotalAmountInEuro = transactions.Sum(t => CurrencyHelper.GetEuroValue(t.Amount, t.TransactionDate));
            decimal calculateTotalFeeInEuro = transactions.Sum(t => CurrencyHelper.GetEuroValue(t.Fee.GetValueOrDefault(), t.TransactionDate));
            decimal calculateTotalCommissionInEuro = transactions.Sum(t => CurrencyHelper.GetEuroValue(t.Commission.GetValueOrDefault(), t.TransactionDate));

            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Справка на транзакции");

            int row = 1;
            int col = 1;
            IXLCell cell;

            // A1
            cell = worksheet.Cell(row, col);
            cell.Value = "Номер на транзакцията";
            cell.Style.Font.Bold = true;

            if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
            {
                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Сума, BGN";
                cell.Style.Font.Bold = true;

                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Такса, BGN";
                cell.Style.Font.Bold = true;

                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Комисионна, BGN";
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

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Такса, EUR";
                    cell.Style.Font.Bold = true;

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Такса, BGN";
                    cell.Style.Font.Bold = true;

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Комисионна, EUR";
                    cell.Style.Font.Bold = true;

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Комисионна, BGN";
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

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Такса, BGN";
                    cell.Style.Font.Bold = true;

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Такса, EUR";
                    cell.Style.Font.Bold = true;

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Комисионна, BGN";
                    cell.Style.Font.Bold = true;

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.Value = "Комисионна, EUR";
                    cell.Style.Font.Bold = true;
                }
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
            {
                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Сума, EUR";
                cell.Style.Font.Bold = true;

                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Такса, EUR";
                cell.Style.Font.Bold = true;

                col++;
                cell = worksheet.Cell(row, col);
                cell.Value = "Комисионна, EUR";
                cell.Style.Font.Bold = true;
            }

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дата на транзакцията";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Номер на карта";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Дата на стълмент";
            cell.Style.Font.Bold = true;

            col++;
            cell = worksheet.Cell(row, col);
            cell.Value = "Съобщение от Борика";
            cell.Style.Font.Bold = true;

            // Move to the begining of the second row
            col = 1;
            row = 2;

            for (int i = 0; i < transactions.Count; i++)
            {
                cell = worksheet.Cell(row, col);
                cell.SetValue(transactions[i].Order);

                if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Amount, transactions[i].TransactionDate));

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Fee.GetValueOrDefault(), transactions[i].TransactionDate));

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Commission.GetValueOrDefault(), transactions[i].TransactionDate));
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
                {
                    if (CurrencyHelper.IsEuroTimePeriod)
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Amount, transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Amount, transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Fee.GetValueOrDefault(), transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Fee.GetValueOrDefault(), transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Commission.GetValueOrDefault(), transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Commission.GetValueOrDefault(), transactions[i].TransactionDate));
                    }
                    else
                    {
                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Amount, transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Amount, transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Fee.GetValueOrDefault(), transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Fee.GetValueOrDefault(), transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetBgnValueFormated(transactions[i].Commission.GetValueOrDefault(), transactions[i].TransactionDate));

                        col++;
                        cell = worksheet.Cell(row, col);
                        cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Commission.GetValueOrDefault(), transactions[i].TransactionDate));
                    }
                }
                else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
                {
                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Amount, transactions[i].TransactionDate));

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Fee.GetValueOrDefault(), transactions[i].TransactionDate));

                    col++;
                    cell = worksheet.Cell(row, col);
                    cell.SetValue(CurrencyHelper.GetEuroValueFormated(transactions[i].Commission.GetValueOrDefault(), transactions[i].TransactionDate));
                }

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(Formatter.DateToBgFormatWithoutYearSuffix(transactions[i].TransactionDate));

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(transactions[i].Card);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(transactions[i].SettlementDate != defaultDate ? Formatter.DateToBgFormatWithoutYearSuffix(transactions[i].SettlementDate) : string.Empty);

                col++;
                cell = worksheet.Cell(row, col);
                cell.SetValue(transactions[i].StatusMessage);

                col = 1;
                row++;
            }

            col = 1;
            row += 2;

            if (CurrencyHelper.CurrencyMode == CurrencyMode.BGN)
            {
                worksheet.Cell($"A{row}").Value = "Общо сума за всички транзакции, BGN";
                worksheet.Cell($"C{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalAmountInEuro * CurrencyHelper.BgnToEuroRate));

                worksheet.Cell($"A{row}").Value = "Общо такси за всички транзакции, BGN";
                worksheet.Cell($"C{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalFeeInEuro * CurrencyHelper.BgnToEuroRate));

                worksheet.Cell($"A{row}").Value = "Общо комисионни за всички транзакции, BGN";
                worksheet.Cell($"C{row}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalCommissionInEuro * CurrencyHelper.BgnToEuroRate));
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.Dual)
            {
                if (CurrencyHelper.IsEuroTimePeriod)
                {
                    worksheet.Cell($"A{row}").Value = "Общо сума за всички транзакции, EUR";
                    worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalAmountInEuro));
                    worksheet.Cell($"A{row}").Value = "Общо сума за всички транзакции, BGN";
                    worksheet.Cell($"C{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalAmountInEuro * CurrencyHelper.BgnToEuroRate));

                    worksheet.Cell($"A{row}").Value = "Общо такси за всички транзакции, EUR";
                    worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalFeeInEuro));
                    worksheet.Cell($"A{row}").Value = "Общо такси за всички транзакции, BGN";
                    worksheet.Cell($"C{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalFeeInEuro * CurrencyHelper.BgnToEuroRate));

                    worksheet.Cell($"A{row}").Value = "Общо комисионни за всички транзакции, EUR";
                    worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalCommissionInEuro));
                    worksheet.Cell($"A{row}").Value = "Общо комисионни за всички транзакции, BGN";
                    worksheet.Cell($"C{row}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalCommissionInEuro * CurrencyHelper.BgnToEuroRate));
                }
                else
                {
                    worksheet.Cell($"A{row}").Value = "Общо сума за всички транзакции, BGN";
                    worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalAmountInEuro * CurrencyHelper.BgnToEuroRate));
                    worksheet.Cell($"A{row}").Value = "Общо сума за всички транзакции, EUR";
                    worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalAmountInEuro));

                    worksheet.Cell($"A{row}").Value = "Общо такси за всички транзакции, BGN";
                    worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalFeeInEuro * CurrencyHelper.BgnToEuroRate));
                    worksheet.Cell($"A{row}").Value = "Общо такси за всички транзакции, EUR";
                    worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalFeeInEuro));

                    worksheet.Cell($"A{row}").Value = "Общо комисионни за всички транзакции, BGN";
                    worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalCommissionInEuro * CurrencyHelper.BgnToEuroRate));
                    worksheet.Cell($"A{row}").Value = "Общо комисионни за всички транзакции, EUR";
                    worksheet.Cell($"B{row}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalCommissionInEuro));
                }
            }
            else if (CurrencyHelper.CurrencyMode == CurrencyMode.EUR)
            {
                worksheet.Cell($"A{row}").Value = "Общо сума за всички транзакции, EUR";
                worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalAmountInEuro));

                worksheet.Cell($"A{row}").Value = "Общо такси за всички транзакции, EUR";
                worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalFeeInEuro));

                worksheet.Cell($"A{row}").Value = "Общо комисионни за всички транзакции, EUR";
                worksheet.Cell($"B{row++}").SetValue<string>(Formatter.DecimalToTwoDecimalPlacesFormat(calculateTotalCommissionInEuro));
            }

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
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Справка на транзакции през ЦВПОС.xlsx");
        }

        public virtual async Task<ActionResult> DownloadPdf(TransactionSearchDO searchDO)
        {
            DateTime? dateFrom = Parser.BgFormatDateStringToDateTime(searchDO.TtDateFrom);
            DateTime? dateTo = Parser.BgFormatDateStringToDateTime(searchDO.TtDateTo);

            int? transactionStatus;

            if (searchDO.TtTransactionStatus.HasValue)
            {
                transactionStatus = (int)searchDO.TtTransactionStatus.Value;
            }
            else
            {
                transactionStatus = null;
            }

            List<BoricaTransactionVO> transactions = await this.EquationControlsRepository.GetBoricaTransactions(
                dateFrom,
                dateTo,
                transactionStatus,
                Enum.GetName(searchDO.TtSortBy.GetType(), searchDO.TtSortBy),
                searchDO.TtSortDesc,
                AppSettings.EPaymentsCommon_MaxPdfExportEntriesCount);

            decimal calculateTotalAmount = transactions.Sum(t => CurrencyHelper.GetEuroValue(t.Amount, t.TransactionDate));
            decimal calculateTotalFee = transactions.Sum(t => CurrencyHelper.GetEuroValue(t.Fee.GetValueOrDefault(), t.TransactionDate));
            decimal calculateTotalCommission = transactions.Sum(t => CurrencyHelper.GetEuroValue(t.Commission.GetValueOrDefault(), t.TransactionDate));

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._TransactionsPdf,
                new TransactionPdfVM()
                {
                    Transactions = transactions,
                    CalculateTotalAmountInEuro = calculateTotalAmount,
                    CalculateTotalFeeInEuro = calculateTotalFee,
                    CalculateTotalCommissionInEuro = calculateTotalCommission
                });

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = "Справка на транзакции през ЦВПОС" + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }

        public virtual async Task<ActionResult> DownloadTransactionPaymentsExcel(int transactionId)
        {
            if (transactionId <= 0)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Транзакцията не е намерена.";

                return this.RedirectToAction(MVC.Transaction.ActionNames.List, MVC.Transaction.Name);
            }

            BoricaTransactionVO boricaTransaction = await this.EquationControlsRepository.GetBoricaTransaction(transactionId);

            if (boricaTransaction == null)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Транзакцията не е намерена.";

                return this.RedirectToAction(MVC.Transaction.ActionNames.List, MVC.Transaction.Name);
            }

            List<PaymentRequestVO> paymentRequests = await this.
                EquationControlsRepository.GetTransactionPaymentRequests(transactionId);

            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Справка транзакция " + boricaTransaction.Order);

            DateTime defaultDate = new DateTime();

            worksheet.Cell("A1").Value = "Номер на транзакцията";

            if (CurrencyHelper.IsEuroTimePeriod)
            {
                worksheet.Cell("B1").Value = "Сума, EUR";
                worksheet.Cell("C1").Value = "Сума, BGN";
                worksheet.Cell("D1").Value = "Такса, EUR";
                worksheet.Cell("E1").Value = "Такса, BGN";
                worksheet.Cell("F1").Value = "Комисионна, EUR";
                worksheet.Cell("G1").Value = "Комисионна, BGN";
            }
            else
            {
                worksheet.Cell("B1").Value = "Сума, BGN";
                worksheet.Cell("C1").Value = "Сума, EUR";
                worksheet.Cell("D1").Value = "Такса, BGN";
                worksheet.Cell("E1").Value = "Такса, EUR";
                worksheet.Cell("F1").Value = "Комисионна, BGN";
                worksheet.Cell("G1").Value = "Комисионна, EUR";
            }

            worksheet.Cell("H1").Value = "Дата на транзакцията";
            worksheet.Cell("I1").Value = "Номер на карта";
            worksheet.Cell("J1").Value = "Дата на стълмент";
            worksheet.Cell("K1").Value = "Съобщение от Борика";

            worksheet.Cell("A1").Style.Font.Bold = true;
            worksheet.Cell("B1").Style.Font.Bold = true;
            worksheet.Cell("C1").Style.Font.Bold = true;
            worksheet.Cell("D1").Style.Font.Bold = true;
            worksheet.Cell("E1").Style.Font.Bold = true;
            worksheet.Cell("F1").Style.Font.Bold = true;
            worksheet.Cell("G1").Style.Font.Bold = true;
            worksheet.Cell("H1").Style.Font.Bold = true;
            worksheet.Cell("I1").Style.Font.Bold = true;
            worksheet.Cell("J1").Style.Font.Bold = true;
            worksheet.Cell("K1").Style.Font.Bold = true;

            worksheet.Cell("A2").SetValue<string>(boricaTransaction.Order);

            if (CurrencyHelper.IsEuroTimePeriod)
            {
                worksheet.Cell("B2").Value = CurrencyHelper.GetEuroValueFormated(boricaTransaction.Amount, boricaTransaction.TransactionDate);
                worksheet.Cell("C2").Value = CurrencyHelper.GetBgnValueFormated(boricaTransaction.Amount, boricaTransaction.TransactionDate);
                worksheet.Cell("D2").Value = boricaTransaction.Fee.HasValue ? CurrencyHelper.GetEuroValueFormated(boricaTransaction.Fee.Value, boricaTransaction.TransactionDate) : "";
                worksheet.Cell("E2").Value = boricaTransaction.Fee.HasValue ? CurrencyHelper.GetBgnValueFormated(boricaTransaction.Fee.Value, boricaTransaction.TransactionDate) : "";
                worksheet.Cell("F2").Value = boricaTransaction.Commission.HasValue ? CurrencyHelper.GetEuroValueFormated(boricaTransaction.Commission.Value, boricaTransaction.TransactionDate) : "";
                worksheet.Cell("G2").Value = boricaTransaction.Commission.HasValue ? CurrencyHelper.GetBgnValueFormated(boricaTransaction.Commission.Value, boricaTransaction.TransactionDate) : "";
            }
            else
            {
                worksheet.Cell("B2").Value = CurrencyHelper.GetBgnValueFormated(boricaTransaction.Amount, boricaTransaction.TransactionDate);
                worksheet.Cell("C2").Value = CurrencyHelper.GetEuroValueFormated(boricaTransaction.Amount, boricaTransaction.TransactionDate);
                worksheet.Cell("D2").Value = boricaTransaction.Fee.HasValue ? CurrencyHelper.GetBgnValueFormated(boricaTransaction.Fee.Value, boricaTransaction.TransactionDate) : "";
                worksheet.Cell("E2").Value = boricaTransaction.Fee.HasValue ? CurrencyHelper.GetEuroValueFormated(boricaTransaction.Fee.Value, boricaTransaction.TransactionDate) : "";
                worksheet.Cell("F2").Value = boricaTransaction.Commission.HasValue ? CurrencyHelper.GetBgnValueFormated(boricaTransaction.Commission.Value, boricaTransaction.TransactionDate) : "";
                worksheet.Cell("G2").Value = boricaTransaction.Commission.HasValue ? CurrencyHelper.GetEuroValueFormated(boricaTransaction.Commission.Value, boricaTransaction.TransactionDate) : "";
            }

            worksheet.Cell("H2").Value = Formatter.DateToBgFormatWithoutYearSuffix(boricaTransaction.TransactionDate);
            worksheet.Cell("I2").Value = boricaTransaction.Card;
            worksheet.Cell("J2").Value = boricaTransaction.SettlementDate != defaultDate ? Formatter.DateToBgFormatWithoutYearSuffix(boricaTransaction.SettlementDate) : "";
            worksheet.Cell("K2").Value = boricaTransaction.StatusMessage;

            worksheet.Cell("A4").Value = "Номер на задължението";
            worksheet.Cell("B4").Value = "Референтен номер на задължението";
            worksheet.Cell("C4").Value = "Дата на създаване";
            worksheet.Cell("D4").Value = "Дата на изтичане";
            worksheet.Cell("E4").Value = "Задължено лице";
            worksheet.Cell("F4").Value = "Основание за плащане";
            if (CurrencyHelper.IsEuroTimePeriod)
            {
                worksheet.Cell("G4").Value = "Сума, EUR";
                worksheet.Cell("H4").Value = "Сума, BGN";
            }
            else
            {
                worksheet.Cell("G4").Value = "Сума, BGN";
                worksheet.Cell("H4").Value = "Сума, EUR";
            }
            worksheet.Cell("I4").Value = "Заявител";
            worksheet.Cell("J4").Value = "Статус на плащането";
            worksheet.Cell("K4").Value = "Статус на задължението";

            worksheet.Cell("A4").Style.Font.Bold = true;
            worksheet.Cell("B4").Style.Font.Bold = true;
            worksheet.Cell("C4").Style.Font.Bold = true;
            worksheet.Cell("D4").Style.Font.Bold = true;
            worksheet.Cell("E4").Style.Font.Bold = true;
            worksheet.Cell("F4").Style.Font.Bold = true;
            worksheet.Cell("G4").Style.Font.Bold = true;
            worksheet.Cell("H4").Style.Font.Bold = true;
            worksheet.Cell("I4").Style.Font.Bold = true;
            worksheet.Cell("J4").Style.Font.Bold = true;
            worksheet.Cell("K4").Style.Font.Bold = true;

            for (int i = 0; i < paymentRequests.Count; i++)
            {
                int row = i + 5;
                var pr = paymentRequests[i];

                worksheet.Cell($"A{row}").Value = pr.PaymentRequestIdentifier;
                worksheet.Cell($"B{row}").Value = pr.PaymentReferenceNumber;
                worksheet.Cell($"C{row}").Value = Formatter.DateToBgFormatWithoutYearSuffix(pr.CreateDate);
                worksheet.Cell($"D{row}").Value = Formatter.DateToBgFormatWithoutYearSuffix(pr.ExpirationDate);
                worksheet.Cell($"E{row}").Value = pr.ApplicantName;
                worksheet.Cell($"F{row}").Value = pr.PaymentReason;

                if (CurrencyHelper.IsEuroTimePeriod)
                {
                    worksheet.Cell($"G{row}").Value = CurrencyHelper.GetEuroValueFormated(pr.PaymentAmount, pr.CreateDate);
                    worksheet.Cell($"H{row}").Value = CurrencyHelper.GetBgnValueFormated(pr.PaymentAmount, pr.CreateDate);
                }
                else
                {
                    worksheet.Cell($"G{row}").Value = CurrencyHelper.GetBgnValueFormated(pr.PaymentAmount, pr.CreateDate);
                    worksheet.Cell($"H{row}").Value = CurrencyHelper.GetEuroValueFormated(pr.PaymentAmount, pr.CreateDate);
                }

                worksheet.Cell($"I{row}").Value = pr.ServiceProviderName;
                worksheet.Cell($"J{row}").Value = Formatter.EnumToDescriptionString(pr.PaymentRequestStatusId);
                worksheet.Cell($"K{row}").Value = pr.ObligationStatusId != null
                    ? Formatter.EnumToDescriptionString(pr.ObligationStatusId)
                    : "Няма стойност";
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
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                string.Format("Справка на транзакция {0}.xlsx", boricaTransaction.Order));
        }

        public virtual async Task<ActionResult> DownloadTransactionPaymentsPdf(int transactionId)
        {
            if (transactionId <= 0)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Транзакцията не е намерена.";

                return this.RedirectToAction(MVC.Transaction.ActionNames.List, MVC.Transaction.Name);
            }

            BoricaTransactionVO boricaTransaction = await this.EquationControlsRepository.GetBoricaTransaction(transactionId);

            if (boricaTransaction == null)
            {
                TempData[Common.TempDataKeys.ErrorMessage] = "Транзакцията не е намерена.";

                return this.RedirectToAction(MVC.Transaction.ActionNames.List, MVC.Transaction.Name);
            }

            List<PaymentRequestVO> paymentRequests = await this.
                EquationControlsRepository.GetTransactionPaymentRequests(transactionId);

            string htmlContent = RenderHelper.RenderHtmlByMvcView(MVC.Shared.Views._PrintTransactionPaymentsPdf,
                new TransactionWithPaymentsVM()
                {
                    BoricaTransaction = boricaTransaction,
                    Payments = paymentRequests
                });

            byte[] data = RenderHelper.RenderPdf(MVC.Shared.Views._PrintPdf, htmlContent);

            string fileName = string.Format("Справка заявки за плащане на транзакция {0}", boricaTransaction.Order) + MimeTypeFileExtension.GetFileExtenstionByMimeType(MimeTypeFileExtension.MIME_APPLICATION_PDF);

            return File(data, MimeTypeFileExtension.MIME_APPLICATION_PDF, fileName);
        }
    }
}