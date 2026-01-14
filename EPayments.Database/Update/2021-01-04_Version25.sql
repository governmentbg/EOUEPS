BEGIN TRANSACTION

BEGIN TRY

ALTER TABLE [dbo].[BoricaTransactions]
	ADD [SettlementDate] DATETIME2 NULL,
		[AuthorizationCode] NVARCHAR(6) NULL,
		[DistributionDate] DATETIME2 NULL,
		[ProductCategory] NVARCHAR(100) NULL,
		[AreaOfIssue] NVARCHAR(100) NULL,
		[IsPaymentSuccessful] BIT NULL,
		[JobCheckResultFailedAttempts] INT NULL,
		[JobCheckResultFailedAttemptsErrors] NVARCHAR(100) NULL,
		[JobLastCheckResultDate] DATETIME2 NULL,
		[JobLastCheckResultTransactionInfo] NVARCHAR(100) NULL;

UPDATE [dbo].[GlobalValues]
	SET Value= '25'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2021-01-04_Version25', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2021-01-04_Version25 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO