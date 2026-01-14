namespace EPayments.Distributions.Interfaces
{
    public interface IDistributionFactory
    {
        IDistributionRevenueCreatable DistributionRevenueCreator();

        ICreateBnbModel BnbModelCreator();

        IBnbXmlDocumentCreator BnbXmlDocumentCreator();
    }
}
