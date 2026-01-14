BEGIN TRANSACTION

BEGIN TRY

ALTER TABLE [dbo].[EserviceClients]
	ADD [IsBoricaVposEnabled] BIT NOT NULL DEFAULT(0) WITH VALUES;

CREATE TABLE [dbo].[DistributionErrors]
(
	[DistributionErrorId]					INT				NOT NULL IDENTITY,
	[Error]									NVARCHAR(500)	NOT NULL,
	[CreatedAd]								DATETIME2		NOT NULL,
	[DistributionRevenueId]					INT				NOT NULL,
	CONSTRAINT [PK_DistributionErrorId]						PRIMARY KEY ([DistributionErrorId]),
	CONSTRAINT [FK_DistributionErrors_DistributionRevenues] FOREIGN KEY (DistributionRevenueId) REFERENCES [dbo].[DistributionRevenues](DistributionRevenueId)
);

CREATE TABLE [dbo].[BoricaTransactionStatuses]
(
	[BoricaTransactionStatusId]					INT				NOT NULL IDENTITY,
	[StatusText]								NVARCHAR(50)	NOT NULL,
	[Alias]										NVARCHAR(100)	NOT NULL,
	CONSTRAINT [PK_BoricaTransactionStatus]						PRIMARY KEY ([BoricaTransactionStatusId])
);

SET IDENTITY_INSERT [dbo].[BoricaTransactionStatuses] ON;

INSERT INTO [dbo].[BoricaTransactionStatuses]
	([BoricaTransactionStatusId], [StatusText], [Alias]	)
VALUES
	(1, 'Pending', 'Очаква плащане'),
	(2, 'Paid', 'Платена'),
	(3, 'TaxReceived', 'Получена такса'),
	(4, 'Distributed', 'Разпределена'),
	(5, 'Canceled', 'Отказана');

SET IDENTITY_INSERT [dbo].[BoricaTransactionStatuses] OFF;

ALTER TABLE [dbo].[BoricaTransactions]
	ADD [Terminal] NVARCHAR(8) NULL,
		[Action] NVARCHAR(4) NULL,
		[Rc] NVARCHAR(4) NULL,
		[Approval] NVARCHAR(6) NULL,
		[Rrn] NVARCHAR(12) NULL,
		[IntRef] NVARCHAR(32) NULL,
		[StatusMessage] NVARCHAR(255) NULL,
		[Card] NVARCHAR(20) NULL,
		[Eci] NVARCHAR(2) NULL,
		[ParesStatus] NVARCHAR(20) NULL,
		CONSTRAINT [FK_BoricaTransactions_BoricaTransactionStatuses] FOREIGN KEY ([TransactionStatusId]) REFERENCES [dbo].[BoricaTransactionStatuses]([BoricaTransactionStatusId]);

UPDATE [dbo].[GlobalValues]
	SET Value= '21'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-12-15_Version21', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-12-15_Version21 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO