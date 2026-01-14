BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '4';
DECLARE @ScriptName NVARCHAR(MAX) = '2020-12-08_Version4';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

DROP TABLE IF EXISTS BoricaDistributionJobLogs;

PRINT 'DistributionJobLogs'

CREATE TABLE [dbo].[DistributionJobLogs](
    [DistributionJobId]                   INT              NOT NULL IDENTITY,
    [Level]                                     NVARCHAR (50)    NULL,
    [Message]                                   NVARCHAR (MAX)   NULL,
    [LogDate]                                   DATETIME         NULL,
    CONSTRAINT [PK_DistributionJobLogs]   PRIMARY KEY ([DistributionJobId])
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