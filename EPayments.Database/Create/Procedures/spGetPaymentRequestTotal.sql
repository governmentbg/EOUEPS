PRINT 'spGetPaymentRequestTotal'
GO

CREATE PROCEDURE [dbo].[spGetPaymentRequestTotal]
AS
BEGIN

SELECT ISNULL(SUM(ISNULL(Counter,0)),0)
	FROM PaymentRequestIdentifiers  
 
END

GO
