CREATE TABLE [tblTranslation] (
  [Language] VARCHAR (10),
  [StringID] LONG ,
  [Translation] LONGTEXT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Language], [StringID])
)
