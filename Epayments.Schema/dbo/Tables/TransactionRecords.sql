CREATE TABLE [dbo].[TransactionRecords](
	[TransactionRecordId] [int] IDENTITY(1,1) NOT NULL,
	[TransactionFileId] [int] NOT NULL,
	[TransactionDate] [datetime2](7) NULL,
	[TransactionAccountingDate] [datetime2](7) NULL,
	[TransactionAmount] [decimal](18, 4) NULL,
	[TransactionReferenceId] [nvarchar](100) NULL,
	[InfoSystemTransactionType] [nvarchar](100) NULL,
	[InfoSystemTransactionDesc] [nvarchar](max) NULL,
	[InfoPaymentType] [nvarchar](100) NULL,
	[InfoDocumentType] [nvarchar](100) NULL,
	[InfoDocumentNumber] [nvarchar](100) NULL,
	[InfoDocumentDate] [datetime2](7) NULL,
	[InfoDocumentNumberDate] [nvarchar](150) NULL,
	[InfoPaymentPeriodBegining] [datetime2](7) NULL,
	[InfoPaymentPeriodEnd] [datetime2](7) NULL,
	[InfoDebtorBulstat] [nvarchar](100) NULL,
	[InfoDebtorEgn] [nvarchar](100) NULL,
	[InfoDebtorLnch] [nvarchar](100) NULL,
	[InfoDebtorName] [nvarchar](300) NULL,
	[InfoDebtorBulstatEgnLnch] [nvarchar](300) NULL,
	[InfoDebtorBulstatEgnLnchName] [nvarchar](600) NULL,
	[InfoAC1AuthorizationCode] [nvarchar](100) NULL,
	[InfoAC1BankCardInfo] [nvarchar](100) NULL,
	[InfoPaymentDetailsRaw] [nvarchar](max) NULL,
	[InfoPaymentReason] [nvarchar](max) NULL,
	[InfoSenderBic] [nvarchar](100) NULL,
	[InfoSenderIban] [nvarchar](100) NULL,
	[InfoSenderName] [nvarchar](max) NULL,
	[InfoSenderIbanName] [nvarchar](max) NULL,
	[TransactionRecordPaymentMethodId] [int] NOT NULL,
	[TransactionRecordRefStatusId] [int] NOT NULL,
	[PaymentRequestId] [int] NULL,
 CONSTRAINT [PK_TransactionRecords] PRIMARY KEY CLUSTERED 
(
	[TransactionRecordId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[TransactionRecords]  WITH CHECK ADD  CONSTRAINT [FK_TransactionRecords_PaymentRequests] FOREIGN KEY([PaymentRequestId])
REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId])
GO

ALTER TABLE [dbo].[TransactionRecords] CHECK CONSTRAINT [FK_TransactionRecords_PaymentRequests]
GO
ALTER TABLE [dbo].[TransactionRecords]  WITH CHECK ADD  CONSTRAINT [FK_TransactionRecords_TransactionFiles] FOREIGN KEY([TransactionFileId])
REFERENCES [dbo].[TransactionFiles] ([TransactionFileId])
GO

ALTER TABLE [dbo].[TransactionRecords] CHECK CONSTRAINT [FK_TransactionRecords_TransactionFiles]
GO
ALTER TABLE [dbo].[TransactionRecords]  WITH CHECK ADD  CONSTRAINT [FK_TransactionRecords_TransactionRecordPaymentMethods] FOREIGN KEY([TransactionRecordPaymentMethodId])
REFERENCES [dbo].[TransactionRecordPaymentMethods] ([TransactionRecordPaymentMethodId])
GO

ALTER TABLE [dbo].[TransactionRecords] CHECK CONSTRAINT [FK_TransactionRecords_TransactionRecordPaymentMethods]
GO
ALTER TABLE [dbo].[TransactionRecords]  WITH CHECK ADD  CONSTRAINT [FK_TransactionRecords_TransactionRecordRefStatuses] FOREIGN KEY([TransactionRecordRefStatusId])
REFERENCES [dbo].[TransactionRecordRefStatuses] ([TransactionRecordRefStatusId])
GO

ALTER TABLE [dbo].[TransactionRecords] CHECK CONSTRAINT [FK_TransactionRecords_TransactionRecordRefStatuses]