CREATE TABLE [dbo].[Emails](
	[EmailId] [int] IDENTITY(1,1) NOT NULL,
	[Recipient] [nvarchar](100) NOT NULL,
	[MailTemplateName] [nvarchar](100) NOT NULL,
	[Context] [nvarchar](max) NULL,
	[NotificationStatusId] [int] NOT NULL,
	[FailedAttempts] [int] NOT NULL,
	[FailedAttemptsErrors] [nvarchar](max) NULL,
	[CreateDate] [datetime2](7) NOT NULL,
	[ModifyDate] [datetime2](7) NOT NULL,
	[Version] [timestamp] NOT NULL,
 CONSTRAINT [PK_Emails] PRIMARY KEY CLUSTERED 
(
	[EmailId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[Emails]  WITH CHECK ADD  CONSTRAINT [FK_Emails_NotificationStatuses] FOREIGN KEY([NotificationStatusId])
REFERENCES [dbo].[NotificationStatuses] ([NotificationStatusId])
GO

ALTER TABLE [dbo].[Emails] CHECK CONSTRAINT [FK_Emails_NotificationStatuses]