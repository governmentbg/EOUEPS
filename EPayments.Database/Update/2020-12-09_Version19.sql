BEGIN TRANSACTION

BEGIN TRY

CREATE TABLE [dbo].[DistributionRevenues]
(
	 [DistributionRevenueId]									INT								NOT NULL IDENTITY,
	 [CreatedAt]												DATETIME2						NOT NULL,
	 [DistributedDate]											DATETIME2						NULL,
	 [IsDistributed]											BIT								NOT NULL DEFAULT(0),
	 [TotalSum]													DECIMAL(18, 4)					NOT NULL,
	 [IsSended]													BIT								NOT NULL DEFAULT(0),
	 [DistributionTypeId]										INT								NOT NULL,
	 CONSTRAINT [PK_DistributionRevenues]														PRIMARY KEY ([DistributionRevenueId]),
	 CONSTRAINT [FK_DistributionRevenues_DistributionTypes]										FOREIGN KEY ([DistributionTypeId]) REFERENCES [dbo].[DistributionTypes](DistributionTypeId)
);

CREATE TABLE [dbo].[DistributionRevenuePayments]
(
	[DistributionRevenuePaymentId]									INT			NOT NULL,
	[EserviceClientId]												INT			NOT NULL,
	[DistributionRevenueId]											INT			NOT NULL,
	CONSTRAINT [PK_DistributionRevenuePayments]									PRIMARY KEY ([DistributionRevenuePaymentId]),
	CONSTRAINT [FK_DistributionRevenuePayments_EserviceClients]					FOREIGN KEY ([EserviceClientId]) REFERENCES [dbo].[EserviceClients]([EserviceClientId]),
	CONSTRAINT [FK_DistributionRevenuePayments_PaymentRequests]					FOREIGN KEY ([DistributionRevenuePaymentId]) REFERENCES [dbo].[PaymentRequests]([PaymentRequestId]),
	CONSTRAINT [FK_DistributionRevenuePayments_DistributionRevenues]			FOREIGN KEY ([DistributionRevenueId]) REFERENCES [dbo].[DistributionRevenues]([DistributionRevenueId])
);

UPDATE [dbo].[GlobalValues]
	SET Value= '19'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-12-09_Version19', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-12-09_Version19 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO