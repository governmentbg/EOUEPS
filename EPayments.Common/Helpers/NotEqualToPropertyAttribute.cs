using System;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace EPayments.Common.Helpers
{
    [AttributeUsage(AttributeTargets.Property)]
    public class NotEqualToPropertyAttribute : ValidationAttribute
    {
        private const string DefaultErrorMessage = "{0} is equal to {1}";

        private readonly string _anotherPropertyName;

        public NotEqualToPropertyAttribute(string anotherPropertyName)
        {
            if (string.IsNullOrWhiteSpace(anotherPropertyName))
            {
                throw new ArgumentNullException(nameof(anotherPropertyName));
            }

            this._anotherPropertyName = anotherPropertyName;
        }

        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            Type type = validationContext.ObjectType;

            PropertyInfo property = type.GetProperty(this._anotherPropertyName);

            if (property == null)
            {
                throw new InvalidCastException(string.Format("Property does not exist in {0}.", type.Name));
            }

            object propertyValue = property.GetValue(validationContext.ObjectInstance);

            if (value != null)
            {
                if (value.GetType() != property.PropertyType || !value.Equals(propertyValue))
                {
                    return ValidationResult.Success;
                }
            }
            else
            {
                if (propertyValue != null)
                {
                    return ValidationResult.Success;
                }
            }

            if (!string.IsNullOrWhiteSpace(this.ErrorMessage))
            {
                return new ValidationResult(String.Format(
                System.Threading.Thread.CurrentThread.CurrentCulture,
                this.ErrorMessage,
                validationContext.DisplayName ?? validationContext.MemberName,
                this._anotherPropertyName));
            }

            return new ValidationResult(String.Format(
                System.Threading.Thread.CurrentThread.CurrentCulture,
                DefaultErrorMessage,
                validationContext.DisplayName ?? validationContext.MemberName,
                this._anotherPropertyName));
        }
    }
}
