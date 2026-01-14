using EPayments.Common;
using EPayments.Common.Data;
using EPayments.Distributions.Interfaces;
using System;

namespace EPayments.Distributions.Implementations
{
    public class DistributionFactory : IDistributionFactory
    {
        private readonly IUnitOfWork UnitOfWork;
        private int BoricaTransactionsToTake;

        public DistributionFactory(IUnitOfWork unitOfWork)
        {
            this.UnitOfWork = unitOfWork ?? throw new ArgumentNullException("unitOfWork is null");
            BoricaTransactionsToTake = AppSettings.EPaymentsJobHost_DistributionTransactionsToTake ?? 1000;
        }

        public ICreateBnbModel BnbModelCreator()
        {
            return new CreateBnbModel();
        }

        public IBnbXmlDocumentCreator BnbXmlDocumentCreator()
        {
            return new BnbXmlDocumentCreator();
        }

        public IDistributionRevenueCreatable DistributionRevenueCreator()
        {
            return new DistributionRevenueCreator(
                this.UnitOfWork, 
                BoricaTransactionsToTake);
        }
    }
}
