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
using EPayments.Data.ViewObjects.Web;
using EPayments.Common.Linq;
using EPayments.Common.Helpers;

namespace EPayments.Data.Repositories.Implementations
{
    internal class SystemRepository : BaseRepository, ISystemRepository
    {
        public SystemRepository(IUnitOfWork unitOfWork)
            : base(unitOfWork)
        {
        }

        public EserviceClient GetEserviceClientById(int eserviceClientId)
        {
            return this.unitOfWork.DbContext.Set<EserviceClient>()
                .SingleOrDefault(e => e.EserviceClientId == eserviceClientId);
        }

        public EserviceClient GetEserviceClientByClientId(string clientId)
        {
            return this.unitOfWork.DbContext.Set<EserviceClient>()
                .SingleOrDefault(e => e.ClientId == clientId);
        }

        public EserviceClient GetEserviceClientByGid(Guid gid)
        {
            return this.unitOfWork.DbContext.Set<EserviceClient>()
                .SingleOrDefault(e => e.Gid == gid);
        }

        public EbankingClient GetEbankingClientByClientId(string clientId)
        {
            return this.unitOfWork.DbContext.Set<EbankingClient>()
                .SingleOrDefault(e => e.ClientId == clientId);
        }

        public List<EserviceClient> GetAllEserviceClients()
        {
            return this.unitOfWork.DbContext.Set<EserviceClient>()
                .ToList();
        }

        public List<Department> GetAllDepartments()
        {
            return this.unitOfWork.DbContext.Set<Department>()
                .ToList();
        }

        public User GetUserByUin(string uin)
        {
            return this.unitOfWork.DbContext.Set<User>()
                .SingleOrDefault(e => e.Egn == uin);
        }

        public User GetUserById(int userId)
        {
            return this.unitOfWork.DbContext.Set<User>()
                .SingleOrDefault(e => e.UserId == userId);
        }

        public EserviceAdminUser GetActiveEserviceAdminUsername(string username)
        {
            return this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .SingleOrDefault(e => e.Username == username && e.IsActive);
        }

        public EserviceAdminUser GetActiveEserviceAdminUserById(int eserviceAdminUserId)
        {
            return this.unitOfWork.DbContext.Set<EserviceAdminUser>()
                .SingleOrDefault(e => e.EserviceAdminUserId == eserviceAdminUserId && e.IsActive);
        }

        public Certificate GetCertificateByThumbprint(string thumbprint)
        {
            return this.unitOfWork.DbContext.Set<Certificate>()
                .SingleOrDefault(e => e.CertificateThumbprint == thumbprint);
        }

        public VposRedirect GetVposRedirectById(int vposRedirectId)
        {
            return this.unitOfWork.DbContext.Set<VposRedirect>()
                .SingleOrDefault(e => e.VposRedirectId == vposRedirectId);
        }

        public VposRedirect GetVposRedirectByGid(Guid gid)
        {
            return this.unitOfWork.DbContext.Set<VposRedirect>()
                .SingleOrDefault(e => e.Gid == gid);
        }

        public bool IsAuthPassLoginExist(Guid gid)
        {
            return this.unitOfWork.DbContext.Set<AuthPassLogin>()
                .Where(e => e.Gid == gid)
                .Any();
        }

        public int GetPaymentRequestCounter(DateTime date)
        {
            var dateParam = new System.Data.SqlClient.SqlParameter
            {
                ParameterName = "Date",
                Value = date
            };

            return this.unitOfWork.DbContext.Database.SqlQuery<int>("spGetPaymentRequestCounter @Date", dateParam).FirstOrDefault();
        }

        public int GetPaymentRequestTotal()
        {
            var total = this.unitOfWork.DbContext.Database.SqlQuery<int>("spGetPaymentRequestTotal").FirstOrDefault();

            while (total > 999999)
            {
                total -= 999999;
            }
            return total;
        }

        public Email GetEmailById(int emailId)
        {
            return this.unitOfWork.DbContext.Set<Email>()
                .SingleOrDefault(e => e.EmailId == emailId);
        }

        public SystemStatsVO GetSystemStats(int? periodYear, int? departmentId, int? eserviceClientId)
        {
            var predicate = PredicateBuilder.True<PaymentRequest>();

            if (periodYear.HasValue)
            {
                var firstDayFromYear = new DateTime(periodYear.Value, 1, 1).Date;
                var firstDayNextYear = new DateTime(periodYear.Value + 1, 1, 1).Date;

                predicate = predicate.And(e => e.CreateDate >= firstDayFromYear && e.CreateDate < firstDayNextYear);
            }

            if (departmentId.HasValue)
            {
                predicate = predicate.And(e => e.EserviceClient.DepartmentId == departmentId.Value);
            }

            if (eserviceClientId.HasValue)
            {
                predicate = predicate.And(e => e.EserviceClientId == eserviceClientId.Value);
            }

            var registeredRequestQuery =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate);

            var registeredRequests = registeredRequestQuery.Count();
            var registeredRequestsAmountInEuro = registeredRequests > 0 ? registeredRequestQuery.Sum(e => CurrencyHelper.GetEuroValue(e.PaymentAmount, e.CreateDate)) : 0;

            var paidViaVposQuery =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .Where(e => e.PaymentRequestStatusId == PaymentRequestStatus.Authorized);

            var paidViaVpos = paidViaVposQuery.Count();
            var paidViaVposAmountInEuro = paidViaVpos > 0 ? paidViaVposQuery.Sum(e => CurrencyHelper.GetEuroValue(e.PaymentAmount, e.CreateDate)) : 0;

            var paidViaBankOrderQuery =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .Where(e => (e.PaymentRequestStatusId == PaymentRequestStatus.Ordered) || (e.PaymentRequestStatusId == PaymentRequestStatus.Paid && !e.IsVposAuthorized));

            var paidViaBankOrder = paidViaBankOrderQuery.Count();
            var paidViaBankOrderAmountInEuro = paidViaBankOrder > 0 ? paidViaBankOrderQuery.Sum(e => CurrencyHelper.GetEuroValue(e.PaymentAmount, e.CreateDate)) : 0;

            var canceledByUser =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .Where(e => e.PaymentRequestStatusId == PaymentRequestStatus.Canceled)
                .Count();

            var pendingRequests =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .Where(e => e.PaymentRequestStatusId == PaymentRequestStatus.Pending)
                .Count();

            var requestsWithAccessCode =
                this.unitOfWork.DbContext.Set<PaymentRequest>()
                .Where(predicate)
                .Where(e => e.PaymentRequestAccessCode != null && e.PaymentRequestAccessCode != "")
                .Count();

            return new SystemStatsVO()
            {
                RegisteredRequests = registeredRequests,
                PaidViaVpos = paidViaVpos,
                PaidViaBankOrder = paidViaBankOrder,
                CanceledByUser = canceledByUser,
                PendingRequests = pendingRequests,
                RequestsWithAccessCode = requestsWithAccessCode,

                RegisteredRequestsAmountInEuro = registeredRequestsAmountInEuro,
                PaidViaVposAmountInEuro = paidViaVposAmountInEuro,
                PaidViaBankOrderAmountInEuro = paidViaBankOrderAmountInEuro,
            };
        }

        public GlobalValue GetGlobalValueByKey(GlobalValueKey key)
        {
            string keyString = key.ToString();

            return this.unitOfWork.DbContext.Set<GlobalValue>()
                .SingleOrDefault(e => e.Key == keyString);
        }

        public InternalAdminUser GetInternalAdminUserByEgn(string egn)
        {
            return this.unitOfWork.DbContext.Set<InternalAdminUser>()
                .SingleOrDefault(e => e.Egn == egn);
        }
    }
}
