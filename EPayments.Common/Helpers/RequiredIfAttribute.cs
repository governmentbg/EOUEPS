using System;
using System.ComponentModel.DataAnnotations;
using System.Linq.Expressions;

namespace EPayments.Common.Helpers
{
    public class RequiredIfAttribute : ValidationAttribute
    {
        private readonly string _condition;

        public RequiredIfAttribute(string condition)
        {
            _condition = condition;
        }

        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            Delegate conditionFunction = CreateExpression(validationContext.ObjectType, _condition);

            bool conditionMet = (bool)conditionFunction.DynamicInvoke(validationContext.ObjectInstance);

            if (conditionMet)
            {
                if (value == null)
                {
                    return new ValidationResult(FormatErrorMessage(null));
                }
            }

            return null;
        }
        private Delegate CreateExpression(Type objectType, string expression)
        {
            LambdaExpression lambdaExpression =
                     System.Linq.Dynamic.DynamicExpression.ParseLambda(objectType, typeof(bool), expression);

            Delegate func = lambdaExpression.Compile();

            return func;
        }
    }
}
