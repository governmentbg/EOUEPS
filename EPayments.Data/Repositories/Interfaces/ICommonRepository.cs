using EPayments.Data.ViewObjects;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface ICommonRepository : IBaseRepository
    {
        PaymentRequest FindPaymentRequestByGuidAndUin(Guid gid, string uin);

        PaymentRequest FindPaymentRequestByIdentifier(string paymentRequestIdentifier);

        string GeneratePaymentRequestAccessCode();
    }
}
