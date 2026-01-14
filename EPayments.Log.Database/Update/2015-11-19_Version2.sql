BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '2';
DECLARE @ScriptName NVARCHAR(MAX) = '2015-11-19_Version2';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------


PRINT 'Create GlobalValues'

CREATE TABLE [dbo].[GlobalValues](
    [GlobalValueId]     INT             IDENTITY (1, 1) NOT NULL,
    [Key]               NVARCHAR (200)  NOT NULL,
    [Value]             NVARCHAR (MAX)  NULL,
    [ModifyDate]        DATETIME2       NOT NULL,

 CONSTRAINT [PK_GlobalValues] PRIMARY KEY ([GlobalValueId])
);


PRINT 'Insert [GlobalValues]'

INSERT [dbo].[GlobalValues] ([Key], [Value], [ModifyDate])
VALUES (N'DatabaseVersion', N'Initial', GETDATE())

CREATE TABLE [dbo].[DatabaseUpdates](
    [DatabaseUpdateId]     INT             IDENTITY (1, 1) NOT NULL,
    [ScriptName]           NVARCHAR (200)  NOT NULL,
    [Notes]                NVARCHAR (MAX)  NULL,
    [UpdateDate]           DATETIME2       NOT NULL,

 CONSTRAINT [PK_DatabaseUpdates] PRIMARY KEY ([DatabaseUpdateId])
);


--------------------------------------------------------------
--Update database version
Update [dbo].[GlobalValues] 
SET [Value] = @DbVersion,
    [ModifyDate] = GETDATE()
WHERE [Key]='DatabaseVersion'
--Add DatabaseUpdate record
INSERT INTO [dbo].[DatabaseUpdates] (ScriptName, Notes, UpdateDate)
VALUES (@ScriptName, @Notes, GETDATE())
-------------------------------------------------------------
--End commands block
COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH






