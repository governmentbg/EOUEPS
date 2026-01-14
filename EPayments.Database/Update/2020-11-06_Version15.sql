BEGIN TRANSACTION
BEGIN TRY

CREATE TABLE [dbo].[DistributionTypes] (
	[DistributionTypeId]					INT					IDENTITY,
	[Name]									NVARCHAR(50)		NOT NULL,
	CONSTRAINT [PK_DistributionTypes]							PRIMARY KEY ([DistributionTypeId])
);

SET IDENTITY_INSERT [dbo].[DistributionTypes] ON

INSERT INTO [dbo].[DistributionTypes]
	([DistributionTypeId], [Name])
VALUES
	(1, 'С банков превод');

SET IDENTITY_INSERT [dbo].[DistributionTypes] OFF

ALTER TABLE [dbo].[Departments]
	ADD [UniqueIdentificationNumber]						NVARCHAR(16)		NULL,
		[UnifiedBudgetClassifier]							NVARCHAR(16)		NULL,
		[IsActive]											BIT					DEFAULT(1) WITH VALUES;

ALTER TABLE [dbo].[EserviceClients]
	ADD [AggregateToParent]									BIT					NOT NULL DEFAULT(0) WITH VALUES,
		[DistributionTypeId]								INT					NOT NULL DEFAULT(1) WITH VALUES,
		[ParentId]											INT					NULL,
		CONSTRAINT [FK_EserviceClients_EserviceClients]							FOREIGN KEY ([ParentId]) REFERENCES [dbo].[EserviceClients] ([EserviceClientId]),
		CONSTRAINT [FK_EserviceClients_DistributionTypes]						FOREIGN KEY ([DistributionTypeId]) REFERENCES [dbo].[DistributionTypes] ([DistributionTypeId])

UPDATE [dbo].[GlobalValues]
	SET Value= '15' 
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-11-06_Version15', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-11-06_Version15 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO