CREATE TABLE [dbo].[LoginAttemptLogs](
	[LoginAttemptLogId] [int] IDENTITY(1,1) NOT NULL,
	[LogDate] [datetime2](7) NOT NULL,
	[IP] [nvarchar](50) NOT NULL,
	[CertificateFile] [varbinary](max) NULL,
	[CertificateIssuer] [nvarchar](max) NULL,
	[CertificatePolicies] [nvarchar](max) NULL,
	[CertificateSubject] [nvarchar](max) NULL,
	[AlternativeSubject] [nvarchar](max) NULL,
	[CertificateThumbprint] [nvarchar](200) NULL,
	[Egn] [nvarchar](50) NULL,
	[NameLatin] [nvarchar](250) NULL,
	[ErrorCode] [nvarchar](30) NULL,
	[IsIisErrorOccurred] [bit] NOT NULL,
	[IsUesParsed] [bit] NOT NULL,
	[IsPersonal] [bit] NULL,
	[IsEgnParsed] [bit] NULL,
	[IsNameParsed] [bit] NULL,
	[IsLoginSuccessful] [bit] NOT NULL,
 CONSTRAINT [PK_LoginAttemptLogs] PRIMARY KEY CLUSTERED 
(
	[LoginAttemptLogId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]