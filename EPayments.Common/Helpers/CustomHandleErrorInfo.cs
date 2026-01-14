using System;
using System.Web.Mvc;

namespace EPayments.Common.Helpers
{
    public class CustomHandleErrorInfo : HandleErrorInfo
    {
        public string ErrorCode { get; set; }

        public int? AttempLogId { get; set; }

        public string Egn { get; set; }

        public string Url { get; set; }

        public string EAuthError { get; set; }

        public CustomHandleErrorInfo(Exception exception, string controllerName = "Home", string actionName = "Error")
            : base(exception, controllerName, actionName)
        {
        }
    }
}
