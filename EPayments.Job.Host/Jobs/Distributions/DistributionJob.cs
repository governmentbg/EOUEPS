using Autofac.Features.OwnedInstances;
using EPayments.Common;
using EPayments.Common.BoricaHelpers;
using EPayments.Common.Data;
using EPayments.CVPosTransaction.Manager;
using EPayments.Data.Repositories.Interfaces;
using EPayments.Distributions.Interfaces;
using EPayments.Distributions.Models.BNB;
using EPayments.Job.Host.Core;
using EPayments.Model.Models;
using NLog;
using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace EPayments.Job.Host.Jobs.Distributions
{
    public class DistributionJob : IJob
    {
        private bool _isDisposed = false;

        private Func<Owned<DisposableTuple<IUnitOfWork, IDistributionFactory, ICVPosRegisterManager, ISystemRepository, IJobRepository>>> DependencyFactory;

        private object syncRoot = new object();

        public DistributionJob(Func<Owned<DisposableTuple<IUnitOfWork, IDistributionFactory, ICVPosRegisterManager, ISystemRepository, IJobRepository>>> dependencyFactory)
        {
            this.DependencyFactory = dependencyFactory ?? throw new ArgumentNullException("dependencyFactory is null");
        }

        public string Name => nameof(DistributionJob);

        public TimeSpan Period => TimeSpan.FromMinutes(AppSettings.EPaymentsJobHost_DistributionJobPeriodInMinutes);

        public void Action(CancellationToken token)
        {
            try
            {
                DisposableTuple<IUnitOfWork, IDistributionFactory, ICVPosRegisterManager, ISystemRepository, IJobRepository> tuple = this.DependencyFactory().Value;
                IUnitOfWork unitOfWork = tuple.Item1;
                ISystemRepository systemRepository = tuple.Item4;
                IJobRepository jobRepository = tuple.Item5;
                int[] time = AppSettings.EPaymentsJobHost_StartTime.Split(new char[] { ':' })
                    .Select(v => int.Parse(v))
                    .ToArray();

                DateTime notBefore = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, time[0], time[1], 0);

                if (DateTime.Now > notBefore)
                {
                    GlobalValue lastInvocationTime = systemRepository
                        .GetGlobalValueByKey(GlobalValueKey.BoricaDistributionJobLastInvocationTime);

                    if (lastInvocationTime == null || (lastInvocationTime.ModifyDate.Year == DateTime.Now.Year &&
                        lastInvocationTime.ModifyDate.Month == DateTime.Now.Month &&
                        lastInvocationTime.ModifyDate.Day == DateTime.Now.Day))
                    {
                        return;
                    }

                    if (lastInvocationTime != null)
                    {
                        lastInvocationTime.ModifyDate = DateTime.Now;
                    }
                    else
                    {
                        unitOfWork.DbContext.Set<GlobalValue>().Add(new GlobalValue()
                        {
                            Key = GlobalValueKey.BoricaDistributionJobLastInvocationTime.ToString(),
                            Value = null,
                            ModifyDate = DateTime.Now
                        });
                    }

                    unitOfWork.Save();


                    Stopwatch sw = new Stopwatch();

                    sw.Start();

                    this.Distribute(token);

                    sw.Stop();

                    JobLogger.Get(JobName.DistributionJob)
                        .Log(LogLevel.Info, string.Format("Distribution job finished in {0}ms.", sw.ElapsedMilliseconds));
                }
            }
            catch (DbUpdateException ex) 
            {
                JobLogger.Get(JobName.DistributionJob).Log(LogLevel.Error, ex.InnerException?.InnerException?.Message);
            }
            catch (OperationCanceledException ex)
            {
                JobLogger.Get(JobName.DistributionJob).Log(LogLevel.Error, ex.Message);
            }
            catch (Exception ex)
            {
                JobLogger.Get(JobName.DistributionJob).Log(LogLevel.Error, $"{ex.Message}, StackTrace -> {ex.StackTrace}");
            }
        }

        public void Dispose()
        {
            if (!_isDisposed)
            {
                DependencyFactory = null;
                _isDisposed = !_isDisposed;
            }
        }

        private void Distribute(CancellationToken token)
        {
            JobLogger.Get(JobName.DistributionJob).Log(LogLevel.Info, $"Start Distribute -> token:{token.IsCancellationRequested}");

            DisposableTuple<IUnitOfWork, IDistributionFactory, ICVPosRegisterManager, ISystemRepository, IJobRepository> tuple = this.DependencyFactory().Value;

            lock (syncRoot)
            {
                IUnitOfWork unitOfWork = tuple.Item1;
                IDistributionFactory distributionFactory = tuple.Item2;
                ISystemRepository systemRepository = tuple.Item4;
                IJobRepository jobRepository = tuple.Item5;

                string bulstat = AppSettings.EPaymentsJobHost_DistributionBulstat;
                string bicCode = AppSettings.EPaymentsJobHost_DistributionBICCode;
                string eGov = AppSettings.EPaymentsJobHost_DistributionSenderName;
                string iban = AppSettings.EPaymentsJobHost_DistributionIban;
                string directoryPath = AppSettings.EPaymentsJobHost_DistributionXmlDirectory;
                string xsdDirectoryPath = AppSettings.EPaymentsJobHost_SchemasDirectory;
                string xsdName = AppSettings.EPaymentsJobHost_XsdFileName;
                int ditributionsToTake = AppSettings.EPaymentsJobHost_DistributionJobItemsToTake;
                int distributionsToSkip = 0;
                string vpn = AppSettings.EPaymentsJobHost_Vpn;
                string vd = AppSettings.EPaymentsJobHost_Vd;
                string firstDescription = AppSettings.EPaymentsJobHost_FirstDescription;
                string secondDescription = AppSettings.EPaymentsJobHost_SecondDescription;
                var obligationTypeList = unitOfWork.DbContext.Set<ObligationType>().ToList();
                DateTime now = DateTime.Now.ToUniversalTime();

                distributionFactory
                    .DistributionRevenueCreator()
                    .Distribute(token);

                while (true)
                {
                    List<DistributionRevenue> distributionRevenues = jobRepository.GetDistributionsWithNoFiles(ditributionsToTake, distributionsToSkip);

                    JobLogger.Get(JobName.DistributionJob).Log(LogLevel.Info, $"Start Distribute -> distributionRevenues:{distributionRevenues}");

                    if (distributionRevenues == null || distributionRevenues.Count == 0)
                    {
                        break;
                    }

                    distributionsToSkip += ditributionsToTake;

                    foreach (DistributionRevenue distributionRevenue in distributionRevenues)
                    {
                        JobLogger.Get(JobName.DistributionJob).Log(LogLevel.Info, $"Distribute -> foreach ->");

                        if (token.IsCancellationRequested == true)
                        {
                            token.ThrowIfCancellationRequested();
                        }

                        distributionRevenue.DistributionRevenuePayments = jobRepository.GetDistributionRevenuePayments(distributionRevenue.DistributionRevenueId);

                        BnbFile bnbFile = distributionFactory.BnbModelCreator()
                           .Create(distributionRevenue, bulstat, eGov, iban, bicCode, vpn, vd, firstDescription, secondDescription, obligationTypeList);

                        IBnbXmlDocumentCreator bnbXmlDocumentCreator = distributionFactory.BnbXmlDocumentCreator();

                        XDocument document = bnbXmlDocumentCreator.CreateDocument(bnbFile);

                        List<string> errors = bnbXmlDocumentCreator.ValidateDocument(document, xsdDirectoryPath, xsdName);

                        if (errors.Count > 0)
                        {
                            string errorMessage = string.Join(", ", errors);

                            JobLogger.Get(JobName.DistributionJob).Log(LogLevel.Warn, errorMessage);
                        }

                        string fileName = string.Format("{0}-{1}.xml", distributionRevenue.DistributionRevenueId, DateTime.Today.ToString("dd.MM.yyyy г."));
                        string filePath = Path.Combine(directoryPath, fileName);

                        using (StreamWriter streamWriter = new StreamWriter(filePath, false, Encoding.GetEncoding("windows-1251")))
                        {
                            document.Save(streamWriter);
                        }

                        UpdateDistribution(distributionRevenue, fileName);

                        unitOfWork.Save();
                    }
                }
            }
        }

        private DistributionRevenue UpdateDistribution(DistributionRevenue distributionRevenue, string fileName)
        {
            distributionRevenue.DistributedDate = DateTime.Now.ToUniversalTime();
            distributionRevenue.IsFileGenerated = true;
            distributionRevenue.IsDistributed = false;
            distributionRevenue.FileName = fileName;

            return distributionRevenue;
        }
    }
}