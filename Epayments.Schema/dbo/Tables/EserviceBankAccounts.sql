CREATE TABLE [dbo].[EserviceBankAccounts](
	[EserviceBankAccountId] [int] IDENTITY(1,1) NOT NULL,
	[BankId] [int] NOT NULL,
	[Iban] [nvarchar](50) NOT NULL,
	[UploadTransactions] [bit] NOT NULL,
	[TransactionsFilesPathUnread] [nvarchar](max) NULL,
	[TransactionsFilesPathRead] [nvarchar](max) NULL,
	[IsActive] [bit] NOT NULL,
 CONSTRAINT [PK_EserviceBankAccounts] PRIMARY KEY CLUSTERED 
(
	[EserviceBankAccountId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[Iban] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[EserviceBankAccounts]  WITH CHECK ADD  CONSTRAINT [FK_EserviceBankAccounts_Banks] FOREIGN KEY([BankId])
REFERENCES [dbo].[Banks] ([BankId])
GO

ALTER TABLE [dbo].[EserviceBankAccounts] CHECK CONSTRAINT [FK_EserviceBankAccounts_Banks]