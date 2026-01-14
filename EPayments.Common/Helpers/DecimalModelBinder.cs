using System.Web.Mvc;

namespace EPayments.Common
{
    public class DecimalModelBinder : IModelBinder
    {
        private readonly string ErrorMessage;

        public DecimalModelBinder(string errorMessage)
        {
            ErrorMessage = !string.IsNullOrEmpty(errorMessage) ? errorMessage : "{0} е невалидно десетично число.";
        }
        
        public object BindModel(
            ControllerContext controllerContext,
            ModelBindingContext bindingContext)
        {
            ValueProviderResult valueResult = bindingContext
                .ValueProvider
                .GetValue(bindingContext.ModelName);

            string propertyName = bindingContext.ModelMetadata.DisplayName ?? bindingContext.ModelMetadata.PropertyName;

            ModelState modelState = new ModelState { Value = valueResult };

            if (valueResult == null && valueResult.AttemptedValue == null)
            {
                return null;
            }

            decimal decimalValue;
            if (!decimal.TryParse(valueResult.AttemptedValue, out decimalValue))
            {
                if (!decimal.TryParse(valueResult.AttemptedValue.Replace(".", ","), out decimalValue))
                    modelState.Errors.Add(string.Format(this.ErrorMessage, propertyName));
            }

            bindingContext.ModelState.Add(bindingContext.ModelName, modelState);

            return decimalValue;
        }
    }
}
