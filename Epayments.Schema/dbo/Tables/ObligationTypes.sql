CREATE TABLE [dbo].[ObligationTypes](
	[ObligationTypeId] [int] IDENTITY(1,1) NOT NULL,
	[Name] [nvarchar](200) NOT NULL,
	[PaymentTypeCode] [nvarchar](50) NOT NULL,
	[AlgorithmId] [int] NULL,
	[IsActive] [bit] NOT NULL,
 CONSTRAINT [PK_ObligationTypeId] PRIMARY KEY CLUSTERED 
(
	[ObligationTypeId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ObligationTypes] ADD  CONSTRAINT [DF__Obligatio__IsAct__3552E9B6]  DEFAULT ((0)) FOR [IsActive]