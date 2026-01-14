BEGIN TRANSACTION

BEGIN TRY

ALTER TABLE [dbo].[DistributionErrors]
	DROP COLUMN [CreatedAd];

ALTER TABLE [dbo].[DistributionErrors]
	ADD [CreatedAt]								DATETIME2		NOT NULL;

ALTER TABLE [dbo].[DistributionRevenues]
	DROP COLUMN [CreatedAd], [IsSended];

ALTER TABLE [dbo].[DistributionRevenues]
	ADD [CreatedAt]								DATETIME2		NOT NULL,
		[IsFileGenerated]						BIT				NOT NULL			 default(0),
		[FileName]								NVARCHAR(50)	NULL;

UPDATE [dbo].[GlobalValues]
	SET [Value]= '27'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2021-03-02_Version27', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2021-03-02_Version27 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO