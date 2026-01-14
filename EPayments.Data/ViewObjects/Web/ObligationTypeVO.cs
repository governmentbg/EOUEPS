using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.ViewObjects.Web
{
    [Serializable()]
    public class ObligationTypeVO
    {
        public int ObligationTypeId { get; set; }
        public string Name { get; set; }
        public bool IsActive { get; set; }
    }
}
