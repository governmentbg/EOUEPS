CREATE TABLE [dbo].[PaymentRequestXmls](
	[PaymentRequestXmlId] [int] IDENTITY(1,1) NOT NULL,
	[RequestContent] [nvarchar](max) NOT NULL,
	[ResponseContent] [nvarchar](max) NOT NULL,
	[IsRequestAccepted] [bit] NOT NULL,
	[CreateDate] [datetime2](7) NOT NULL,
 CONSTRAINT [PK_PaymentRequestXmls] PRIMARY KEY CLUSTERED 
(
	[PaymentRequestXmlId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]