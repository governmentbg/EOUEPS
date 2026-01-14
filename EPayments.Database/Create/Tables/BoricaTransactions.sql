PRINT 'BoricaTransactions'
GO

CREATE TABLE [dbo].[BoricaTransactions] (
	[BoricaTransactionId]					INT					IDENTITY,
	[Gid]									UNIQUEIDENTIFIER 	NOT NULL,
	[Amount]								DECIMAL(18,4)		NOT NULL,
	[Order]									NVARCHAR(16)		NOT NULL,
	[Description]							NVARCHAR(50)		NULL,
	[Fee]									DECIMAL(18,4)		NULL,
	[Commission]							DECIMAL(18,4)		NULL,
	[TransactionDate]						DATETIME2			NOT NULL,
	[TransactionStatus]						INT					NULL,
	CONSTRAINT [PK_BoricaTransactions]							PRIMARY KEY ([BoricaTransactionId])
);
GO