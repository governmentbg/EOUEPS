PRINT 'TRIGGER trForUpdatePaymentRequest'

GO

create TRIGGER trForUpdatePaymentRequest
ON PaymentRequests
FOR UPDATE
AS

DECLARE @PaymentRequestId int, @OldPaymentRequestStatusId int, @NewPaymentRequestStatusId int;

BEGIN
	SELECT @PaymentRequestId=PaymentRequestId, @NewPaymentRequestStatusId=PaymentRequestStatusId FROM inserted;
	SELECT @OldPaymentRequestStatusId=PaymentRequestStatusId FROM deleted;

		
	if (@NewPaymentRequestStatusId is not null and @OldPaymentRequestStatusId != @NewPaymentRequestStatusId)
	begin
		insert into PaymentRequestStatusLogs (PaymentRequestId, PaymentRequestStatusId, ChangeDate) 
		values (@PaymentRequestId, @NewPaymentRequestStatusId, GETDATE())
	end
END
GO