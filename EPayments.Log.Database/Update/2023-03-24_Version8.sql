BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '31';
DECLARE @ScriptName NVARCHAR(MAX) = '2023-03-24_Version8';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

PRINT 'CVPosTransactionFixJobLogs'

CREATE TABLE [dbo].[CVPosTransactionFixJobLogs](
    [CVPosTransactionFixJobId]                  INT              NOT NULL IDENTITY,
    [Level]                                     NVARCHAR (50)    NULL,
    [Message]                                   NVARCHAR (MAX)   NULL,
    [LogDate]                                   DATETIME         NULL,
    CONSTRAINT [PK_CVPosTransactionFixJobLogs]  PRIMARY KEY ([CVPosTransactionFixJobId])
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

--Update [GlobalValues] table
INSERT INTO [dbo].[GlobalValues] ([Key], ModifyDate)
VALUES ('CVPosTransactionFixJobInvocationTime', DATEADD(day, -1, GETDATE()))
-------------------------------------------------------------
--End commands block
COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH