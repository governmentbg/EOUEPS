BEGIN TRANSACTION

BEGIN TRY

-- Obligation Types

CREATE TABLE [dbo].[ObligationTypes] (
	[ObligationTypeId]					INT				IDENTITY,
	[Name]								NVARCHAR(200)	NOT NULL,
	[IsActive]							BIT				NOT NULL DEFAULT(0), 
	CONSTRAINT [PK_ObligationTypeId]					PRIMARY KEY ([ObligationTypeId])
);

CREATE INDEX IDX_ObligationTypes_Name ON [dbo].[ObligationTypes]([Name]);

SET IDENTITY_INSERT [dbo].[ObligationTypes] ON;

INSERT INTO [dbo].[ObligationTypes]
	([ObligationTypeId], [Name], [IsActive])
VALUES
	(1, 'Общо задължение', 1),
	(2, 'Плащане на ЕАУ', 1),
	(3, 'Плащане на такса за детска градина', 1),
	(4, 'Плащане на глоби към КАТ', 1),
	(5, 'Плащане на местни данъци и такси', 1)

SET IDENTITY_INSERT [dbo].[ObligationTypes] OFF;

-- TransactionRecords

CREATE TABLE [dbo].[BoricaTransactions] (
	[BoricaTransactionId]					INT					IDENTITY,
	[Gid]									UNIQUEIDENTIFIER 	NOT NULL,
	[Amount]								DECIMAL(18,4)		NOT NULL,
	[Order]									NVARCHAR(16)		NOT NULL,
	[Description]							NVARCHAR(50)		NULL,
	[Fee]									DECIMAL(18,4)		NULL,
	[Commission]							DECIMAL(18,4)		NULL,
	[TransactionDate]						DATETIME2			NOT NULL,
	[TransactionStatusId]					INT					NULL,
	CONSTRAINT [PK_BoricaTransactions]							PRIMARY KEY ([BoricaTransactionId])
);

CREATE TABLE [dbo].[BoricaTransactionPaymentRequest] (
	[BoricaTransactionId]													INT			NOT NULL,
	[PaymentRequestId]														INT			NOT NULL,
	CONSTRAINT [PK_BoricaTransactionPaymentRequest]										PRIMARY KEY ([BoricaTransactionId], [PaymentRequestId]),
	CONSTRAINT [FK_BoricaTransactionPaymentRequest_BoricaTransactions]					FOREIGN KEY ([BoricaTransactionId]) REFERENCES [dbo].[BoricaTransactions]([BoricaTransactionId]),
	CONSTRAINT [FK_BoricaTransactionPaymentRequest_PaymentRequests]						FOREIGN KEY ([PaymentRequestId]) REFERENCES [dbo].[PaymentRequests]([PaymentRequestId])
);

-- EserviceClients

ALTER TABLE [dbo].[EserviceClients]
	ADD [ObligationTypeId]									INT			NOT NULL DEFAULT(1) WITH VALUES,
	CONSTRAINT [FK_EserviceClients_ObligationTypes]						FOREIGN KEY ([ObligationTypeId]) REFERENCES [dbo].[ObligationTypes](ObligationTypeId);

-- Updated Database version

UPDATE [dbo].[GlobalValues]
	SET Value= '17'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-11-21_Version17', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-11-21_Version17 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO