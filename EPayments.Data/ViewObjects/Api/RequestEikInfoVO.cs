using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Api
{
    public class RequestEikInfoVO
    {
        public int DepartmentId { get; set; }
        public string Name { get; set; }
        public string UniqueIdentificationNumber { get; set; }
        public bool IsActive { get; set; }
        public int EserviceClientId { get; set; }
        public string AisName { get; set; }
        public string AccountBank { get; set; }
        public string AccountBIC { get; set; }
        public string AccountIBAN { get; set; }
    }
}
