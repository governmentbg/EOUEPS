BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '11';
DECLARE @ScriptName NVARCHAR(MAX) = '2017-01-18_Version11';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------






--update na EserviceClients -> dobavqne na kolonite svarzani s Transaction funkcionalnostta

--update na PaymentRequestStatus

--add na tablicite svarzani s Transaction funkcionalnostta: 
--TransactionFiles
--TransactionRecords
--EserviceAdminUsers
--EserviceBankAccounts
--EserviceAdminUserBankAccounts
--TransactionRecordRefStatuses
--LogsDB:ProcessTransactionFilesJobLogs









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






