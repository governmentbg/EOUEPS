using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Service.Api.Common.CustomExceptions
{
    public class CustomServiceException : Exception
    {
        public CustomServiceException(string message)
            : base(message)
        {
        }
    }
}
