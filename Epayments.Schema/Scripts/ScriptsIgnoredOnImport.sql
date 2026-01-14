
USE [EPayments.Production]
GO

/****** Object:  Table [dbo].[Departments]    Script Date: 30.5.2022 г. 9:05:55 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EserviceClients]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  View [dbo].[vwDepartmentAisClients]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[AuthPassLogins]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[Banks]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[BoricaTransactionPaymentRequest]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[BoricaTransactions]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[BoricaTransactionStatuses]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[Certificates]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[DatabaseUpdates]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[Disclaimers]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[DisclaimerTemplates]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[DistributionErrors]    Script Date: 30.5.2022 г. 9:05:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[DistributionRevenuePayments]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[DistributionRevenues]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[DistributionTypes]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EbankingAccessLogs]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EbankingClients]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[Emails]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EserviceAdminUserBankAccounts]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EserviceAdminUsers]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EserviceBankAccounts]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EserviceDeliveryNotifications]    Script Date: 30.5.2022 г. 9:05:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EserviceNotifications]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[EventRegisterNotifications]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[GlobalValues]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[InternalAdminUsers]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[LoginAttemptLogs]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[NotificationStatuses]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[ObligationStatuses]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[ObligationTypes]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[PaidStatusPaymentMethods]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[PaymentRequestIdentifiers]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[PaymentRequestObligationLogs]    Script Date: 30.5.2022 г. 9:05:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[PaymentRequests]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[PaymentRequestStatuses]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[PaymentRequestStatusLogs]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[PaymentRequestXmls]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[Permissions]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[TransactionFiles]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[TransactionRecordPaymentMethods]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[TransactionRecordRefStatuses]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[TransactionRecords]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[TransactionRequestIdentifiers]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[UinTypes]    Script Date: 30.5.2022 г. 9:05:59 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[Users]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[VposBoricaRequests]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[VposClients]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[VposDskEcommRequests]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[VposEpayRequests]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[VposFiBankRequests]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[VposRedirects]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[VposResults]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  StoredProcedure [dbo].[spGetPaymentRequestCounter]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/****** Object:  StoredProcedure [dbo].[spGetPaymentRequestTotal]    Script Date: 30.5.2022 г. 9:06:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
