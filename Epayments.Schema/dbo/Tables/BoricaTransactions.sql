CREATE TABLE [dbo].[BoricaTransactions](
	[BoricaTransactionId] [int] IDENTITY(1,1) NOT NULL,
	[Gid] [uniqueidentifier] NOT NULL,
	[Amount] [decimal](18, 4) NOT NULL,
	[Order] [nvarchar](16) NOT NULL,
	[Description] [nvarchar](50) NULL,
	[Fee] [decimal](18, 4) NULL,
	[Commission] [decimal](18, 4) NULL,
	[TransactionDate] [datetime2](7) NOT NULL,
	[TransactionStatusId] [int] NULL,
	[Terminal] [nvarchar](8) NULL,
	[Action] [nvarchar](4) NULL,
	[Rc] [nvarchar](4) NULL,
	[Approval] [nvarchar](6) NULL,
	[Rrn] [nvarchar](12) NULL,
	[IntRef] [nvarchar](32) NULL,
	[StatusMessage] [nvarchar](255) NULL,
	[Card] [nvarchar](20) NULL,
	[Eci] [nvarchar](2) NULL,
	[ParesStatus] [nvarchar](20) NULL,
	[SettlementDate] [datetime2](7) NULL,
	[AuthorizationCode] [nvarchar](6) NULL,
	[DistributionDate] [datetime2](7) NULL,
	[ProductCategory] [nvarchar](100) NULL,
	[AreaOfIssue] [nvarchar](100) NULL,
	[IsPaymentSuccessful] [bit] NULL,
	[JobCheckResultFailedAttempts] [int] NULL,
	[JobCheckResultFailedAttemptsErrors] [nvarchar](100) NULL,
	[JobLastCheckResultDate] [datetime2](7) NULL,
	[JobLastCheckResultTransactionInfo] [nvarchar](100) NULL,
 [BoricaTransactionSettlement] NVARCHAR(MAX) NULL, 
    CONSTRAINT [PK_BoricaTransactions] PRIMARY KEY CLUSTERED 
(
	[BoricaTransactionId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[BoricaTransactions]  WITH CHECK ADD  CONSTRAINT [FK_BoricaTransactions_BoricaTransactionStatuses] FOREIGN KEY([TransactionStatusId])
REFERENCES [dbo].[BoricaTransactionStatuses] ([BoricaTransactionStatusId])
GO

ALTER TABLE [dbo].[BoricaTransactions] CHECK CONSTRAINT [FK_BoricaTransactions_BoricaTransactionStatuses]