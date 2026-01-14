using EPayments.Common.Data;
using EPayments.Distributions.Interfaces;
using EPayments.Model.Enums;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.Entity;
using System.Threading;
using EPayments.Common;
using EPayments.Model.DataObjects;
using Newtonsoft.Json;
using EPayments.Common.Helpers;

namespace EPayments.Distributions.Implementations
{
    public class DistributionRevenueCreator : IDistributionRevenueCreatable, IDisposable
    {
        private readonly int _boricaTransactionsToTake;
        private readonly IUnitOfWork _unitOfWork;
        private bool _isDisposed = false;
        private HashSet<EserviceClient> _cachedClients = new HashSet<EserviceClient>();
        private readonly int? _parentClient = AppSettings.EPaymentsJobHost_DistributionJobParentEserviceClient;

        public DistributionRevenueCreator(IUnitOfWork unitOfWork, int boricaTransactionsToTake)
        {
            _unitOfWork = unitOfWork ?? throw new ArgumentNullException("unitOfWork is null");
            _boricaTransactionsToTake = boricaTransactionsToTake;
        }

        public List<DistributionRevenue> Distribute(CancellationToken token)
        {
            List<DistributionRevenue> distributionRevenues = null;

            while (true)
            {
                if (token.IsCancellationRequested)
                {
                    token.ThrowIfCancellationRequested();
                }

                List<BoricaTransaction> boricaTransactions = _unitOfWork.DbContext.Set<BoricaTransaction>()
                    .Include(bt => bt.PaymentRequests)
                    .Where(bt => bt.TransactionStatusId == (int)BoricaTransactionStatusEnum.TaxReceived)
                    .OrderBy(bt => bt.BoricaTransactionId)
                    .Take(_boricaTransactionsToTake)
                    .ToList();

                if (boricaTransactions == null || boricaTransactions.Count == 0)
                {
                    break;
                }

                List<PaymentRequest> paymentRequests = boricaTransactions
                    .SelectMany(bt => bt.PaymentRequests.Where(pr => pr.ObligationStatusId == ObligationStatusEnum.Paid))
                    .ToList();

                if (paymentRequests.Count > 0)
                {
                    if (distributionRevenues == null)
                    {
                        distributionRevenues = new List<DistributionRevenue>();
                    }

                    UpdateClientsCache(paymentRequests);
                    //DateTime now = DateTime.Now.ToUniversalTime();

                    Dictionary<int, HashSet<DistributionRevenuePayment>> distributionsDictionary = CreateDistributionsDictionary(
                        paymentRequests, boricaTransactions);

                    //add new DistributionRevenuePayments to distributionRevenue and update distributionRevenue.TotalSum
                    foreach (int distributionTypeId in distributionsDictionary.Keys)
                    {
                        if (distributionsDictionary[distributionTypeId].Count > 0)
                        {
                            var distributionRevenue = distributionRevenues.FirstOrDefault(dr => dr.DistributionTypeId == distributionTypeId);

                            if (distributionRevenue == null)
                            {
                                distributionRevenue = new DistributionRevenue()
                                {
                                    CreatedAt = DateTime.Now.ToUniversalTime(),
                                    IsFileGenerated = false,
                                    IsDistributed = false,
                                    DistributionTypeId = distributionTypeId
                                };

                                distributionRevenues.Add(distributionRevenue);
                            }

                            foreach (DistributionRevenuePayment distributionRevenuePayment in distributionsDictionary[distributionTypeId])
                            {
                                if (!distributionRevenue.DistributionRevenuePayments.Any(drp => drp.DistributionRevenuePaymentId == distributionRevenuePayment.DistributionRevenuePaymentId))
                                {
                                    distributionRevenue.DistributionRevenuePayments.Add(distributionRevenuePayment);

                                    if (CurrencyHelper.IsEuroTimePeriod)
                                    {
                                        distributionRevenue.TotalSum += CurrencyHelper.GetEuroValue(distributionRevenuePayment.PaymentRequest.PaymentAmount, distributionRevenuePayment.PaymentRequest.CreateDate);
                                    }
                                    else
                                    {
                                        distributionRevenue.TotalSum += CurrencyHelper.GetBgnValue(distributionRevenuePayment.PaymentRequest.PaymentAmount, distributionRevenuePayment.PaymentRequest.CreateDate);
                                    }
                                }
                            }
                        }
                    }
                }

                //update status of transaction and paymentRequests, and add new distribution revenues to DB
                using (var transaction = _unitOfWork.BeginTransaction())
                {
                    if (paymentRequests.Count > 0)
                    {
                        foreach (PaymentRequest paymentRequest in paymentRequests)
                        {
                            paymentRequest.ObligationStatusId = ObligationStatusEnum.ForDistribution;
                        }
                    }

                    foreach (BoricaTransaction boricaTransaction in boricaTransactions)
                    {
                        boricaTransaction.TransactionStatusId = (int)BoricaTransactionStatusEnum.Distributed;
                    }

                    if (distributionRevenues != null && distributionRevenues.Count > 0)
                    {
                        IEnumerable<DistributionRevenue> newDistributionRevenues =
                            distributionRevenues.Where(dr => _unitOfWork.DbContext.Entry(dr).State == EntityState.Detached);

                        if (newDistributionRevenues != null && newDistributionRevenues.Count() > 0)
                        {
                            foreach (DistributionRevenue newDistributionRevenue in newDistributionRevenues)
                            {
                                _unitOfWork.DbContext.Entry(newDistributionRevenue).State = EntityState.Added;
                            }
                        }
                    }

                    _unitOfWork.Save();

                    transaction.Commit();
                }
            }

            return distributionRevenues;
        }

        public void Dispose()
        {
            if (!_isDisposed)
            {
                _cachedClients = null;
                _isDisposed = !_isDisposed;
            }
        }

        private Dictionary<int, HashSet<DistributionRevenuePayment>> CreateDistributionsDictionary(IEnumerable<PaymentRequest> paymentRequests,
                                                                                                   List<BoricaTransaction> boricaTransactions)
        {
            Dictionary<int, HashSet<DistributionRevenuePayment>> distributionsDictionary = new Dictionary<int, HashSet<DistributionRevenuePayment>>();

            foreach (PaymentRequest paymentRequest in paymentRequests)
            {
                var boricaTransaction = boricaTransactions.First(bt => bt.PaymentRequests.Contains(paymentRequest));
                var distributionRevenuePayment = CreateDistributionRevenuePayment(paymentRequest, boricaTransaction);

                if (!distributionsDictionary.TryGetValue(distributionRevenuePayment.EserviceClient.DistributionTypeId,
                    out HashSet<DistributionRevenuePayment> distributionRevenuePayments))
                {
                    distributionRevenuePayments = new HashSet<DistributionRevenuePayment>();
                    distributionsDictionary.Add(distributionRevenuePayment.EserviceClient.DistributionTypeId, distributionRevenuePayments);
                }

                distributionRevenuePayments.Add(distributionRevenuePayment);
            }

            return distributionsDictionary;
        }

        private void UpdateClientsCache(IEnumerable<PaymentRequest> paymentRequests)
        {
            HashSet<int> ids = new HashSet<int>(paymentRequests.Select(pr => pr.EserviceClientId));
            HashSet<string> idsOfRegions = new HashSet<string>();
            if (_parentClient != null)
            {
                string[] paymentRequestsAdditionalInfos = paymentRequests.Where(x => x.EserviceClientId == _parentClient).Select(x => x.AdditionalInformation).ToArray();
                foreach (var paymentRequestsAdditionalInfo in paymentRequestsAdditionalInfos)
                {
                    string sanitizedJson = JsonHelper.RepairBrokenAdditionalInformationJsonString(paymentRequestsAdditionalInfo);

                    var additionalInfoModel = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(sanitizedJson);

                    if (additionalInfoModel.regionClientId != null)
                    {
                        idsOfRegions.Add(additionalInfoModel.regionClientId);
                    }
                }
            }

            IEnumerable<EserviceClient> clients = _unitOfWork.DbContext.Set<EserviceClient>().
                Where(ec => ids.Contains(ec.EserviceClientId) || idsOfRegions.Contains(ec.ClientId));

            _cachedClients.UnionWith(clients);
        }

        private EserviceClient GetEserviceClient(int id, string regionClientId)
        {
            EserviceClient client;
            if (regionClientId == String.Empty)
            {
                client = _cachedClients.FirstOrDefault(c => c.EserviceClientId == id);
            }
            else
            {
                client = _cachedClients.FirstOrDefault(c => c.ClientId == regionClientId);
            }

            if (client != null)
            {
                return client;
            }

            //try get client from DB if not present in cache
            if (regionClientId == String.Empty)
            {
                client = _unitOfWork.DbContext.Set<EserviceClient>().SingleOrDefault(c => c.EserviceClientId == id);
            }
            else
            {
                client = _unitOfWork.DbContext.Set<EserviceClient>().SingleOrDefault(c => c.ClientId == regionClientId);
            }

            if (client == null)
            {
                return null;
            }

            _cachedClients.Add(client);

            return client;
        }

        private DistributionRevenuePayment CreateDistributionRevenuePayment(PaymentRequest paymentRequest, BoricaTransaction boricaTransaction)
        {
            string regionClientId = String.Empty;

            //code added due to separation of Stolicna Obstina to its regions
            if (_parentClient != null && paymentRequest.EserviceClientId == _parentClient)
            {
                string sanitizedJson = JsonHelper.RepairBrokenAdditionalInformationJsonString(paymentRequest.AdditionalInformation);

                var additionalInfoModel = JsonConvert.DeserializeObject<MDT_ExtendedInfoJson>(sanitizedJson);

                if (additionalInfoModel.regionClientId != null)
                {
                    regionClientId = additionalInfoModel.regionClientId;
                }

            }

            EserviceClient currentClient = GetEserviceClient(paymentRequest.EserviceClientId, regionClientId);

            if (currentClient == null)
            {
                throw new InvalidOperationException(string.Format("Eservice client with id {0} was not found.", paymentRequest.EserviceClientId));
            }

            while (true)
            {
                if (currentClient.AggregateToParent == false)
                {
                    return new DistributionRevenuePayment()
                    {
                        EserviceClientId = currentClient.EserviceClientId,
                        EserviceClient = currentClient,
                        DistributionRevenuePaymentId = paymentRequest.PaymentRequestId,
                        PaymentRequest = paymentRequest,
                        BoricaTransactionId = boricaTransaction.BoricaTransactionId,
                        BoricaTransaction = boricaTransaction
                    };
                }
                else
                {
                    //ToDo: This is probably never used!
                    if (currentClient.ParentId == null)
                    {
                        throw new InvalidOperationException(string.Format("Eservice client with id {0} does not have parent.", currentClient.EserviceClientId));
                    }

                    currentClient = GetEserviceClient((int)currentClient.ParentId, String.Empty);

                    if (currentClient == null)
                    {
                        throw new NullReferenceException(string.Format("Eservice client with id {0} does not exist.", currentClient.ParentId));
                    }
                }
            }
        }
    }
}
