using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.CVPosTransaction.DataObjects
{
    public class TransactionPerDateDO
    {
        public TransactionPerDateDO()
        {
        }
        public string @Event { get; set; }
        public string TID { get; set; }
        public DateTime DateEvent { get; set;}
        public string Agency { get; set; }
    }
}

