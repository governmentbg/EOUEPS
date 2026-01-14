CREATE TABLE [dbo].[TransactionFiles](
	[TransactionFileId] [int] IDENTITY(1,1) NOT NULL,
	[EserviceBankAccountId] [int] NOT NULL,
	[FileName] [nvarchar](max) NULL,
	[BankStatementIban] [nvarchar](100) NOT NULL,
	[BankStatementDate] [datetime2](7) NULL,
	[BankStatementNumber] [nvarchar](100) NULL,
	[CreateDate] [datetime2](7) NOT NULL,
 CONSTRAINT [PK_TransactionFiles] PRIMARY KEY CLUSTERED 
(
	[TransactionFileId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[TransactionFiles]  WITH CHECK ADD  CONSTRAINT [FK_TransactionFiles_EserviceBankAccounts] FOREIGN KEY([EserviceBankAccountId])
REFERENCES [dbo].[EserviceBankAccounts] ([EserviceBankAccountId])
GO

ALTER TABLE [dbo].[TransactionFiles] CHECK CONSTRAINT [FK_TransactionFiles_EserviceBankAccounts]