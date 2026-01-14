using EPayments.Data.ViewObjects;
using EPayments.Data.ViewObjects.Web;
using EPayments.Model.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Data.Repositories.Interfaces
{
    public interface ISystemRepository : IBaseRepository
    {
        EserviceClient GetEserviceClientById(int eserviceClientId);

        EserviceClient GetEserviceClientByClientId(string clientId);

        EserviceClient GetEserviceClientByGid(Guid gid);

        EbankingClient GetEbankingClientByClientId(string clientId);

        List<EserviceClient> GetAllEserviceClients();

        List<Department> GetAllDepartments();

        User GetUserByUin(string uin);

        User GetUserById(int userId);

        EserviceAdminUser GetActiveEserviceAdminUsername(string username);

        EserviceAdminUser GetActiveEserviceAdminUserById(int eserviceAdminUserId);

        Certificate GetCertificateByThumbprint(string thumbprint);

        VposRedirect GetVposRedirectById(int vposRedirectId);

        VposRedirect GetVposRedirectByGid(Guid gid);

        bool IsAuthPassLoginExist(Guid gid);

        int GetPaymentRequestCounter(DateTime date);

        int GetPaymentRequestTotal();

        Email GetEmailById(int emailId);

        SystemStatsVO GetSystemStats(int? periodYear, int? departmentId, int? eserviceClientId);

        GlobalValue GetGlobalValueByKey(GlobalValueKey key);

        InternalAdminUser GetInternalAdminUserByEgn(string egn);
    }
}
