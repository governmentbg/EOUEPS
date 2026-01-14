drop table TransactionReportPayments
drop table TransactionReports

ALTER TABLE EserviceClients DROP COLUMN DisableUniqueRequestConstraint

ALTER TABLE PaymentRequests ADD [AisPaymentId] [nvarchar](max) NULL


SET IDENTITY_INSERT [dbo].[Banks] ON 

GO
INSERT [dbo].[Banks] ([BankId], [Name], [BIC], [CertificateId]) VALUES (1, N'Банка ДСК', N'STSABGSF', NULL)
GO
INSERT [dbo].[Banks] ([BankId], [Name], [BIC], [CertificateId]) VALUES (2, N'УниКредит Булбанк', N'UNCRBGSF', NULL)
GO
SET IDENTITY_INSERT [dbo].[Banks] OFF
GO





PRINT 'TransactionRecordPaymentMethods'
GO

CREATE TABLE [dbo].[TransactionRecordPaymentMethods](
    [TransactionRecordPaymentMethodId]  INT             NOT NULL IDENTITY,
    [Alias]                             NVARCHAR(50)    NOT NULL,
    [Name]                              NVARCHAR(100)   NOT NULL,
    [Description]                       NVARCHAR(MAX)   NULL,
 CONSTRAINT [PK_TransactionRecordPaymentMethods] PRIMARY KEY ([TransactionRecordPaymentMethodId])
);

GO


PRINT 'TransactionRecordRefStatuses'
GO

CREATE TABLE [dbo].[TransactionRecordRefStatuses](
    [TransactionRecordRefStatusId]  INT             NOT NULL IDENTITY,
    [Alias]                         NVARCHAR(50)    NOT NULL,
    [Name]                          NVARCHAR(100)   NOT NULL,
    [Description]                   NVARCHAR(MAX)   NULL,
 CONSTRAINT [PK_TransactionRecordRefStatuses] PRIMARY KEY ([TransactionRecordRefStatusId])
);

GO


-------

PRINT 'EserviceBankAccounts'
GO

CREATE TABLE [dbo].[EserviceBankAccounts](
    [EserviceBankAccountId]                 INT                 NOT NULL IDENTITY,
    [BankId]                                INT                 NOT NULL,
    [Iban]                                  NVARCHAR(50)        NOT NULL UNIQUE,
    [UploadTransactions]                    BIT                 NOT NULL,
    [TransactionsFilesPathUnread]           NVARCHAR(MAX)       NULL,
    [TransactionsFilesPathRead]             NVARCHAR(MAX)       NULL,
    [IsActive]                              BIT                 NOT NULL,
CONSTRAINT [PK_EserviceBankAccounts]                        PRIMARY KEY ([EserviceBankAccountId]),
CONSTRAINT [FK_EserviceBankAccounts_Banks]                  FOREIGN KEY ([BankId])     REFERENCES [dbo].[Banks] ([BankId])
);






---------

PRINT 'TransactionFiles'
GO

CREATE TABLE [dbo].[TransactionFiles](
    [TransactionFileId]             INT             NOT NULL IDENTITY,
    [EserviceBankAccountId]         INT             NOT NULL,
    [FileName]                      NVARCHAR(MAX)   NULL,
    [BankStatementIban]             NVARCHAR(100)   NOT NULL,
    [BankStatementDate]             DATETIME2       NULL,
    [BankStatementNumber]           NVARCHAR(100)   NULL,
    [CreateDate]                    DATETIME2       NOT NULL,
CONSTRAINT [PK_TransactionFiles]                           PRIMARY KEY ([TransactionFileId]),
CONSTRAINT [FK_TransactionFiles_EserviceBankAccounts]      FOREIGN KEY ([EserviceBankAccountId])        REFERENCES [dbo].[EserviceBankAccounts] ([EserviceBankAccountId])
);

GO


---------


PRINT 'TransactionRecords'
GO

CREATE TABLE [dbo].[TransactionRecords](
    [TransactionRecordId]                   INT             NOT NULL IDENTITY,
    [TransactionFileId]                     INT             NOT NULL,
    [TransactionDate]                       DATETIME2       NULL,
    [TransactionAccountingDate]             DATETIME2       NULL,
    [TransactionAmount]                     DECIMAL(18, 4)  NULL,
    [TransactionReferenceId]                NVARCHAR(100)   NULL,
    [InfoSystemTransactionType]             NVARCHAR(100)   NULL,
    [InfoSystemTransactionDesc]             NVARCHAR(MAX)   NULL,
    [InfoPaymentType]                       NVARCHAR(100)   NULL,
    [InfoDocumentType]                      NVARCHAR(100)   NULL,
    [InfoDocumentNumber]                    NVARCHAR(100)   NULL,
    [InfoDocumentDate]                      DATETIME2       NULL,
    [InfoDocumentNumberDate]                NVARCHAR(150)   NULL,
    [InfoPaymentPeriodBegining]             DATETIME2       NULL,
    [InfoPaymentPeriodEnd]                  DATETIME2       NULL,
    [InfoDebtorBulstat]                     NVARCHAR(100)   NULL,
    [InfoDebtorEgn]                         NVARCHAR(100)   NULL,
    [InfoDebtorLnch]                        NVARCHAR(100)   NULL,
    [InfoDebtorName]                        NVARCHAR(300)   NULL,
    [InfoDebtorBulstatEgnLnch]              NVARCHAR(300)   NULL,
    [InfoDebtorBulstatEgnLnchName]          NVARCHAR(600)   NULL,
    [InfoAC1AuthorizationCode]              NVARCHAR(100)   NULL,
    [InfoAC1BankCardInfo]                   NVARCHAR(100)   NULL,
    [InfoPaymentDetailsRaw]                 NVARCHAR(MAX)   NULL,
    [InfoPaymentReason]                     NVARCHAR(MAX)   NULL,
    [InfoSenderBic]                         NVARCHAR(100)   NULL,
    [InfoSenderIban]                        NVARCHAR(100)   NULL,
    [InfoSenderName]                        NVARCHAR(MAX)   NULL,
    [InfoSenderIbanName]                    NVARCHAR(MAX)   NULL,
    [TransactionRecordPaymentMethodId]      INT             NOT NULL,
    [TransactionRecordRefStatusId]          INT             NOT NULL,
    [PaymentRequestId]                      INT             NULL,

CONSTRAINT [PK_TransactionRecords]                                      PRIMARY KEY ([TransactionRecordId]),
CONSTRAINT [FK_TransactionRecords_TransactionFiles]                     FOREIGN KEY ([TransactionFileId])                   REFERENCES [dbo].[TransactionFiles] ([TransactionFileId]),
CONSTRAINT [FK_TransactionRecords_TransactionRecordPaymentMethods]      FOREIGN KEY ([TransactionRecordPaymentMethodId])    REFERENCES [dbo].[TransactionRecordPaymentMethods] ([TransactionRecordPaymentMethodId]),
CONSTRAINT [FK_TransactionRecords_TransactionRecordRefStatuses]         FOREIGN KEY ([TransactionRecordRefStatusId])        REFERENCES [dbo].[TransactionRecordRefStatuses] ([TransactionRecordRefStatusId]),
CONSTRAINT [FK_TransactionRecords_PaymentRequests]                      FOREIGN KEY ([PaymentRequestId])                    REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
);

GO


--------

PRINT 'EserviceAdminUsers'
GO

CREATE TABLE [dbo].[EserviceAdminUsers](
    [EserviceAdminUserId]                   INT                 NOT NULL IDENTITY,
    [DepartmentId]                          INT                 NOT NULL,
    [Username]                              NVARCHAR(100)       NOT NULL UNIQUE,
    [PasswordHash]                          NVARCHAR(MAX)       NOT NULL,
    [PasswordSalt]                          NVARCHAR(MAX)       NOT NULL,
    [IpAddressesForAccess]                  NVARCHAR(MAX)       NULL,
    [Name]                                  NVARCHAR(100)       NOT NULL,
    [Email]                                 NVARCHAR(100)       NULL,
    [InsufficientAmountNotifications]       BIT                 NOT NULL,
    [OverpaidAmountNotifications]           BIT                 NOT NULL,
    [ReferencedNotifications]               BIT                 NOT NULL,
    [NotReferencedNotifications]            BIT                 NOT NULL,
    [IsActive]                              BIT                 NOT NULL,
CONSTRAINT [PK_EserviceAdminUsers]                              PRIMARY KEY ([EserviceAdminUserId]),
CONSTRAINT [FK_EserviceAdminUsers_Departments]                  FOREIGN KEY ([DepartmentId])     REFERENCES [dbo].[Departments] ([DepartmentId])
);



------


PRINT 'EserviceAdminUserBankAccounts'
GO

CREATE TABLE [dbo].[EserviceAdminUserBankAccounts](
    [EserviceAdminUserBankAccountId]            INT                 NOT NULL IDENTITY,
    [EserviceAdminUserId]                       INT                 NOT NULL,
    [EserviceBankAccountId]                     INT                 NOT NULL,
    [IsActive]                                  BIT                 NOT NULL,
CONSTRAINT [PK_EserviceAdminUserBankAccounts]                               PRIMARY KEY ([EserviceAdminUserBankAccountId]),
CONSTRAINT [FK_EserviceAdminUserBankAccounts_EserviceAdminUsers]            FOREIGN KEY ([EserviceAdminUserId])     REFERENCES [dbo].[EserviceAdminUsers] ([EserviceAdminUserId]),
CONSTRAINT [FK_EserviceAdminUserBankAccounts_EserviceBankAccounts]          FOREIGN KEY ([EserviceBankAccountId])   REFERENCES [dbo].[EserviceBankAccounts] ([EserviceBankAccountId])
);




--------------

SET IDENTITY_INSERT [dbo].[TransactionRecordPaymentMethods] ON 

GO
INSERT [dbo].[TransactionRecordPaymentMethods] ([TransactionRecordPaymentMethodId], [Alias], [Name], [Description]) VALUES (1, N'BankOrder', N'По банков път', NULL)
GO
INSERT [dbo].[TransactionRecordPaymentMethods] ([TransactionRecordPaymentMethodId], [Alias], [Name], [Description]) VALUES (2, N'POS', N'Физически ПОС', NULL)
GO
INSERT [dbo].[TransactionRecordPaymentMethods] ([TransactionRecordPaymentMethodId], [Alias], [Name], [Description]) VALUES (3, N'VPOS', N'Виртуален ПОС', NULL)
GO
SET IDENTITY_INSERT [dbo].[TransactionRecordPaymentMethods] OFF
GO
SET IDENTITY_INSERT [dbo].[TransactionRecordRefStatuses] ON 

GO
INSERT [dbo].[TransactionRecordRefStatuses] ([TransactionRecordRefStatusId], [Alias], [Name], [Description]) VALUES (1, N'NotReferenced', N'Друго', NULL)
GO
INSERT [dbo].[TransactionRecordRefStatuses] ([TransactionRecordRefStatusId], [Alias], [Name], [Description]) VALUES (2, N'ReferencedSuccessfully', N'Платено задължение', NULL)
GO
INSERT [dbo].[TransactionRecordRefStatuses] ([TransactionRecordRefStatusId], [Alias], [Name], [Description]) VALUES (3, N'ReferencedInsufficientAmount', N'Недостатъчна сума', NULL)
GO
INSERT [dbo].[TransactionRecordRefStatuses] ([TransactionRecordRefStatusId], [Alias], [Name], [Description]) VALUES (4, N'ReferencedOverpaidAmount', N'Надплатена сума', NULL)
GO
SET IDENTITY_INSERT [dbo].[TransactionRecordRefStatuses] OFF
GO



-----------

update PaymentRequestStatuses set Name=N'Платено по банков път' where Alias='Ordered'


SET IDENTITY_INSERT [dbo].[EserviceBankAccounts] ON 
GO
INSERT [dbo].[EserviceBankAccounts] ([EserviceBankAccountId], [BankId], [Iban], [UploadTransactions], [TransactionsFilesPathUnread], [TransactionsFilesPathRead], [IsActive]) 
VALUES (1, 2, N'BG97UNCR96603119995010', 1, N'C:\FTP\unicredit', N'C:\FTP\unicredit_nacid\Read', 1)
GO
SET IDENTITY_INSERT [dbo].[EserviceBankAccounts] OFF
GO
SET IDENTITY_INSERT [dbo].[EserviceAdminUsers] ON 

GO
INSERT [dbo].[EserviceAdminUsers] ([EserviceAdminUserId], [DepartmentId], [Username], [PasswordHash], [PasswordSalt], [IpAddressesForAccess], [Name], [Email], [InsufficientAmountNotifications], [OverpaidAmountNotifications], [ReferencedNotifications], [NotReferencedNotifications], [IsActive]) VALUES (1, 1, N'emil_prod', N'AKCM5mChIKu1ztNnQ9hZ4r9PdIwOdy3HoyTeHHsWsN+1Ot98WuDcuJMl2Qi7l8tm5g==', N'1urqFSoAUT6ldJKSqYFn4A==', NULL, N'Емил Дечев Денчовски', NULL, 0, 0, 0, 0, 1)
GO
SET IDENTITY_INSERT [dbo].[EserviceAdminUsers] OFF
GO
SET IDENTITY_INSERT [dbo].[EserviceAdminUserBankAccounts] ON 

GO
INSERT [dbo].[EserviceAdminUserBankAccounts] ([EserviceAdminUserBankAccountId], [EserviceAdminUserId], [EserviceBankAccountId], [IsActive]) VALUES (1, 1, 1, 1)
GO
SET IDENTITY_INSERT [dbo].[EserviceAdminUserBankAccounts] OFF
GO


----------

update GlobalValues set Value='10' where [key]='DatabaseVersion'

SET IDENTITY_INSERT [dbo].[DatabaseUpdates] ON 
GO
INSERT [dbo].[DatabaseUpdates] ([DatabaseUpdateId], [ScriptName], [Notes], [UpdateDate]) 
VALUES (9, N'2017-02-21_Version10', NULL, CAST(N'2017-02-21 22:15:26.6700000' AS DateTime2))
GO
SET IDENTITY_INSERT [dbo].[DatabaseUpdates] OFF
GO
