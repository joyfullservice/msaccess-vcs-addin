CREATE TABLE [MSysIMEXColumns] (
  [Attributes] LONG ,
  [DataType] SHORT ,
  [FieldName] VARCHAR (64),
  [IndexType] UNSIGNED BYTE ,
  [SkipColumn] BIT ,
  [SpecID] LONG ,
  [Start] SHORT ,
  [Width] SHORT ,
   CONSTRAINT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY (SpecID, FieldName)
)
