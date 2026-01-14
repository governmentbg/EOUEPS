BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '8';
DECLARE @ScriptName NVARCHAR(MAX) = '2016-02-25_Version8';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

delete from emails where MailTemplateName='AccessCodeDeactivatedMessage' and NotificationStatusId != 2


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






