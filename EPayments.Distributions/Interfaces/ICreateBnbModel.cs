using System.Collections.Generic;
using EPayments.Distributions.Models.BNB;
using EPayments.Model.Models;

namespace EPayments.Distributions.Interfaces
{
    public interface ICreateBnbModel
    {
        BnbFile Create(DistributionRevenue distributionRevenue,
            string bulstat,
            string senderName,
            string iban,
            string bicCode,
            string vpn,
            string vd,
            string firstDescription,
            string secondDescription,
            List<ObligationType> oblicationTypeList);
    }
}
