using EPayments.Common.Data;
using EPayments.Data.Repositories.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;
using EPayments.Model.Models;
using EPayments.Data.ViewObjects;
using EPayments.Model.Enums;
using EPayments.Common.Helpers;

namespace EPayments.Data.Repositories.Implementations
{
    internal class CommonRepository : BaseRepository, ICommonRepository
    {
        public CommonRepository(IUnitOfWork unitOfWork)
            : base(unitOfWork)
        {
        }

        public PaymentRequest FindPaymentRequestByGuidAndUin(Guid gid, string uin)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    where pr.Gid == gid && pr.ApplicantUin == uin
                    select pr)
                    .Include(p => p.ObligationType)
                    .SingleOrDefault();
        }

        public PaymentRequest FindPaymentRequestByIdentifier(string paymentRequestIdentifier)
        {
            return (from pr in this.unitOfWork.DbContext.Set<PaymentRequest>()
                    where pr.PaymentRequestIdentifier == paymentRequestIdentifier
                    select pr)
                    .Include(p => p.ObligationType)
                    .SingleOrDefault();
        }

        public string GeneratePaymentRequestAccessCode()
        {
            CodeGenerator codeGenerator = new CodeGenerator();
            codeGenerator.Minimum = 10;
            codeGenerator.Maximum = 10;
            codeGenerator.ConsecutiveCharacters = true;
            codeGenerator.RepeatCharacters = true;

            while (true)
            {
                string code = codeGenerator.Generate();

                if (!this.unitOfWork.DbContext.Set<PaymentRequest>().Any(e => e.PaymentRequestAccessCode == code))
                {
                    return code;
                }
            }
        }
    }
}
