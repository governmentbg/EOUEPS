using EPayments.Common.Helpers;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace EPayments.Web.DataObjects
{
    public class PaymentRequestDO : IValidatableObject
    {
        private const string DateFormat = "dd.MM.yyyy";

        public static Dictionary<int, string> TimeByMinutes = new Dictionary<int, string>()
        {
            { 0, "0:00"},
            { 30, "0:30"},
            { 60, "1:00"},
            { 90, "1:30"},
            { 120, "2:00"},
            { 150, "2:30"},
            { 180, "3:00"},
            { 210, "3:30"},
            { 240, "4:00"},
            { 270, "4:30"},
            { 300, "5:00"},
            { 330, "5:30"},
            { 360, "6:00"},
            { 390, "6:30"},
            { 420, "7:00"},
            { 450, "7:30"},
            { 480, "8:00"},
            { 510, "8:30"},
            { 540, "9:00"},
            { 570, "9:30"},
            { 600, "10:00"},
            { 630, "10:30"},
            { 660, "11:00"},
            { 690, "11:30"},
            { 720, "12:00"},
            { 750, "12:30"},
            { 780, "13:00"},
            { 810, "13:30"},
            { 840, "14:00"},
            { 870, "14:30"},
            { 900, "15:00"},
            { 930, "15:30"},
            { 960, "16:00"},
            { 990, "16:30"},
            { 1020, "17:00"},
            { 1050, "17:30"},
            { 1080, "18:00"},
            { 1110, "18:30"},
            { 1140, "19:00"},
            { 1170, "19:30"},
            { 1200, "20:00"},
            { 1230, "20:30"},
            { 1260, "21:00"},
            { 1290, "21:30"},
            { 1320, "22:00"},
            { 1350, "22:30"},
            { 1380, "23:00"},
            { 1410, "23:30"},
        };

        public string PaymentTypeCode { get; set; }

        [Required(ErrorMessage = "Сумата на задължението не е въведена.")]
        [Range(0.01, 100000, ErrorMessage = "Невалидна сума не може да е по-малка от {1} и по-голяма от {2}.")]
        public decimal PaymentAmount { get; set; }


        [Required(ErrorMessage = "Описанието на задължението е задължително.")]
        [MaxLength(200, ErrorMessage = "Описанието на задължението не може да е по-дълго от {1} символа.")]
        [Display(Name = "Основание за плащане")]
        public string PaymentReason { get; set; }

        [Required(ErrorMessage = "Вида на идентификация на задълженото лице е задължително.")]
        [Display(Name = "Вид на идентификация на задълженото лице")]
        public UinType ApplicantUinTypeId { get; set; }

        [Required(ErrorMessage = "Данните за идентификация на задълженото лице са задължителни.")]
        [MaxLength(50, ErrorMessage = "Данните за идентификация на задълженото лице не може да са по-дълги от {1} символа.")]
        [Display(Name = "Данни за идентификация на задълженото лице")]
        public string ApplicantUin { get; set; }

        [Required(ErrorMessage = "Видa на документa е задължителен.")]
        [MaxLength(50, ErrorMessage = "Видa на документа не може да е по-дълът от {1} символа.")]
        [Display(Name = "Вид документ")]
        public string PaymentReferenceType { get; set; }

        [Required(ErrorMessage = "Референтен номер на плащането е задължителен.")]
        [MaxLength(50, ErrorMessage = "Референтен номер на плащането не може да е по-дълъг от {1} символа.")]
        [Display(Name = "Номер на документа (референтен номер ORN)")]
        public string PaymentReferenceNumber { get; set; }

        [Required(ErrorMessage = "Името на задълженото лице е задължително.")]
        [MaxLength(500, ErrorMessage = "Името на задълженото лице не може да е по-дълго от {1} символа.")]
        [Display(Name = "Задължено лице")]
        public string ApplicantName { get; set; }

        public DateTime ExpirationDate { get; set; } = DateTime.Now.AddDays(1);

        [Display(Name = "Час")]
        [Required(ErrorMessage = "Името на задълженото лице е задължително.")]
        public int Minutes { get; set; } = 720;

        [Required]
        [Display(Name = "Дата, на която изтича задължението")]
        [ValidateStringIsDate(RegexValidator = DateFormat, ErrorMessage = "Невалидна дата, на която изтича задължението.")]
        public string ExpirationDateAsString
        {
            get
            {
                return this.ExpirationDate.ToString(DateFormat);
            }
            set
            {
                DateTime date;

                if (DateTime.TryParseExact(value, DateFormat,
                    System.Threading.Thread.CurrentThread.CurrentCulture,
                    System.Globalization.DateTimeStyles.None,
                    out date))
                {
                    int minutes = 0;

                    if (PaymentRequestDO.TimeByMinutes.ContainsKey(this.Minutes))
                    {
                        minutes = this.Minutes;
                    }
                    
                    this.ExpirationDate = date.AddMinutes(minutes);
                }
            }
        }

        [MaxLength(500, ErrorMessage = "Допълнителна информация не може да е по-дълга от {1} символа.")]
        [Display(Name = "Допълнителна информация")]
        public string AdditionalInformation { get; set; }

        [Required(ErrorMessage = "Видa на задължението е задължителен.")]
        [Display(Name = "Вид на задължението")]
        public string ObligationType { get; set; }

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if (this.ExpirationDate < DateTime.Now)
            {
                yield return new ValidationResult("Дата, на която изтича задължението трябва да е по голяма от днешната дата.");
            }

            if (!PaymentRequestDO.TimeByMinutes.ContainsKey(this.Minutes))
            {
                yield return new ValidationResult("Избрана е невалидна опция за час.");
            }
            
            switch (this.ApplicantUinTypeId)
            {
                case UinType.Egn:
                    string invalidIDMessage = "Невалидни данни на задълженото лице за ЕГН.";

                    if (!Regex.IsMatch(ApplicantUin, @"\d{2}([024][1-9]|[135][0-2])([0][1-9]|[12][0-9]|[3][01])\d{4}"))
                    {
                        yield return new ValidationResult(invalidIDMessage);
                    }

                    EgnHelper egnHelper = new EgnHelper(ApplicantUin);

                    if (!egnHelper.IsValid())
                    {
                        yield return new ValidationResult(invalidIDMessage);
                    }
                    break;
                //case UinType.Lnch:
                //    if (!Regex.IsMatch(ApplicantUin, @"\d{2}([024][1-9]|[135][0-2])([0][1-9]|[12][0-9]|[3][01])\d{4}"))
                //    {
                //        yield return new ValidationResult("Невалидни данни на задълженото лице за личен номер на чужденец.");
                //    }
                //    break;
                case UinType.Bulstat:
                    if (ApplicantUin.Length != 9 && ApplicantUin.Length != 13)
                    {
                        yield return new ValidationResult("Булстата е с дължина различна от 9 или 13 символа.");
                    }

                    if (!Regex.IsMatch(ApplicantUin, @"^[0-9]+$"))
                    {
                        yield return new ValidationResult("Булстата може да съдържа само цифри.");
                    }

                    break;
            }
        }

        public static List<ObligationType> ObligationTypesList { get; set; } = new List<ObligationType>();

    }
}