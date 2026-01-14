BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '5';
DECLARE @ScriptName NVARCHAR(MAX) = '2020-12-22_Version5';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

PRINT 'CVPosTransactionJobLogs'

CREATE TABLE [dbo].[CVPosTransactionJobLogs](
    [CVPosTransactionJobId]                     INT              NOT NULL IDENTITY,
    [Level]                                     NVARCHAR (50)    NULL,
    [Message]                                   NVARCHAR (MAX)   NULL,
    [LogDate]                                   DATETIME         NULL,
    CONSTRAINT [PK_CVPosTransactionJobLogs]   PRIMARY KEY ([CVPosTransactionJobId])
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