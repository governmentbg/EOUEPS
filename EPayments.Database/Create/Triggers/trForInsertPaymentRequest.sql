PRINT 'TRIGGER trForInsertPaymentRequest'

GO

create TRIGGER trForInsertPaymentRequest
ON PaymentRequests
FOR INSERT
AS

DECLARE @PaymentRequestId int, @PaymentRequestStatusId int;

BEGIN
	SELECT @PaymentRequestId=PaymentRequestId, @PaymentRequestStatusId=PaymentRequestStatusId FROM inserted;
		
	insert into PaymentRequestStatusLogs (PaymentRequestId, PaymentRequestStatusId, ChangeDate) 
	values (@PaymentRequestId, @PaymentRequestStatusId, GETDATE())
	
END
GO