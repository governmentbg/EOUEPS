-- update triggers

CREATE TRIGGER trForInsertPaymentRequest
ON PaymentRequests
FOR INSERT
AS

DECLARE @InsertedPaymentRequestId int, @PaymentRequestStatusId int, @ObligationStatusId int, @InsertChangeDate DATETIME2 = GETDATE();

BEGIN
	SELECT @InsertedPaymentRequestId=PaymentRequestId, @PaymentRequestStatusId=PaymentRequestStatusId, @ObligationStatusId = ObligationStatusId FROM inserted;
		
	insert into PaymentRequestStatusLogs (PaymentRequestId, PaymentRequestStatusId, ChangeDate) 
	values (@InsertedPaymentRequestId, @PaymentRequestStatusId, @InsertChangeDate)

	IF (@ObligationStatusId IS NOT NULL)
	BEGIN
		INSERT INTO PaymentRequestObligationLogs (PaymentRequestId, ObligationStatusId, ChangeDate)
		VALUES (@InsertedPaymentRequestId, @ObligationStatusId, @InsertChangeDate)
	END
	
END
GO

CREATE TRIGGER trForUpdatePaymentRequest
ON PaymentRequests
FOR UPDATE
AS

DECLARE @PaymentRequestId int, @OldPaymentRequestStatusId int, @NewPaymentRequestStatusId int, @OldObligationStatusId int, @NewObligationStatusId int, @ChangeDate DATETIME2 = GETDATE();

BEGIN
	SELECT @PaymentRequestId=PaymentRequestId, @NewPaymentRequestStatusId=PaymentRequestStatusId, @NewObligationStatusId = ObligationStatusId FROM inserted;
	SELECT @OldPaymentRequestStatusId=PaymentRequestStatusId, @OldObligationStatusId = ObligationStatusId FROM deleted;

		
	if (@NewPaymentRequestStatusId is not null and @OldPaymentRequestStatusId != @NewPaymentRequestStatusId)
	begin
		insert into PaymentRequestStatusLogs (PaymentRequestId, PaymentRequestStatusId, ChangeDate) 
		values (@PaymentRequestId, @NewPaymentRequestStatusId, @ChangeDate)
	end

	if (@NewObligationStatusId IS NOT NULL AND @OldObligationStatusId != @NewObligationStatusId)
	BEGIN
		INSERT INTO PaymentRequestObligationLogs (PaymentRequestId, ObligationStatusId, ChangeDate)
		VALUES (@PaymentRequestId, @NewObligationStatusId, @ChangeDate)
	END
END
GO

-- update database values
UPDATE [dbo].[GlobalValues]
	SET Value= '14' 
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-11-03_Version14', NULL, GETDATE());
GO