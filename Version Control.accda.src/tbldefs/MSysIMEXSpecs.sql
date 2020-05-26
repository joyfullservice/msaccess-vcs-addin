CREATE TABLE [MSysIMEXSpecs] (
  [DateDelim] VARCHAR (2),
  [DateFourDigitYear] BIT ,
  [DateLeadingZeros] BIT ,
  [DateOrder] SHORT ,
  [DecimalPoint] VARCHAR (2),
  [FieldSeparator] VARCHAR (2),
  [FileType] SHORT ,
  [SpecID] AUTOINCREMENT,
  [SpecName] VARCHAR (64) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SpecType] UNSIGNED BYTE ,
  [StartRow] LONG ,
  [TextDelim] VARCHAR (2),
  [TimeDelim] VARCHAR (2)
)
