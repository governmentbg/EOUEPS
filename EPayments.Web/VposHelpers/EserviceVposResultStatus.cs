using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPayments.Web.VposHelpers
{
    public enum EserviceVposResultStatus
    {
        Success,
        Failure,
        CanceledByUser,
        RequestIsCanceled,
        RequestIsExpired,
        RequestIsPaid
    }
}