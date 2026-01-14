BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '9';
DECLARE @ScriptName NVARCHAR(MAX) = '2016-03-23_Version9';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

CREATE TABLE [dbo].[EbankingAccessLogs](
    [EbankingAccessLogId]                   INT                 NOT NULL IDENTITY,
    [EbankingClientId]                      INT                 NOT NULL,
    [PaymentRequestId]                      INT                 NOT NULL,
    [IpAddress]                             NVARCHAR (50)       NOT NULL,
    [AccessDate]                            DATETIME2           NOT NULL,
CONSTRAINT [PK_EbankingAccessLogs]                             PRIMARY KEY ([EbankingAccessLogId]),
CONSTRAINT [FK_EbankingAccessLogs_EbankingClients]             FOREIGN KEY ([EbankingClientId])            REFERENCES [dbo].[EbankingClients] ([EbankingClientId]),
CONSTRAINT [FK_EbankingAccessLogs_PaymentRequests]             FOREIGN KEY ([PaymentRequestId])            REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
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






