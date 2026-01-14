BEGIN TRANSACTION

BEGIN TRY

ALTER TABLE [dbo].[InternalAdminUsers]
	ADD [Permissions] INT NULL;

CREATE TABLE [dbo].[Permissions] (
	[PermissionId]								INT					NOT NULL IDENTITY,
	[Alias]										NVARCHAR(50)		NULL,
	[PermissionText]							NVARCHAR(50)		NULL,
	[Description]								NVARCHAR(200)		NULL,
	CONSTRAINT [PK_Permissions_PermissionId]						PRIMARY KEY ([PermissionId])
);

INSERT INTO [dbo].[Permissions]
	([Alias], [PermissionText])
VALUES
	('ViewReferences', 'Преглед на справки'),
	('DistributionReferences', 'Преглед на справки за разпределение'),
	('Modify', 'Управление');

UPDATE [dbo].[GlobalValues]
	SET Value= '18'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-12-02_Version18', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-12-02_Version18 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO