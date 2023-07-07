CREATE TABLE [tblTableData] (
  [TableIcon] VARCHAR (2),
  [TableName] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [FormatType] LONG ,
  [Flags] LONG ,
  [IsSystem] BIT ,
  [IsHidden] BIT ,
  [IsLocal] BIT ,
  [IsOther] BIT 
)
