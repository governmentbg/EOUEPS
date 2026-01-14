using System;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Threading;

namespace EPayments.Common.Helpers
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ValidateStringIsDateAttribute : ValidationAttribute
    {
        public string RegexValidator { get; set; }

        public string MinDateError { get; set; } = "The date {0} must exceed today date";

        public bool GreaterThanToday { get; set; }

        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            if (string.IsNullOrWhiteSpace(RegexValidator))
            {
                return ValidationResult.Success;
            }

            if (value == null)
            {
                return ValidationResult.Success;
            }

            DateTime date;

            if (DateTime.TryParseExact(value.ToString(),
                this.RegexValidator,
                Thread.CurrentThread.CurrentCulture,
                DateTimeStyles.None,
                out date))
            {
                if (GreaterThanToday == true && 
                    (date.Year <= DateTime.Now.Year && date.Month <= DateTime.Now.Month && date.Day <= DateTime.Now.Day))
                {
                    return new ValidationResult(String.Format(
                        Thread.CurrentThread.CurrentCulture,
                        this.MinDateError,
                        (validationContext.DisplayName ?? validationContext.MemberName)));
                }

                return ValidationResult.Success;
            }

            return new ValidationResult(String.Format(
                Thread.CurrentThread.CurrentCulture,
                this.ErrorMessage,
                (validationContext.DisplayName ?? validationContext.MemberName)));
        }
    }
}
