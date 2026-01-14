using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Log.ActionLogger
{
    public interface IActionLogger
    {
        void LogAction(string clientId, object postData, object responseData);
    }
}
