CREATE TABLE [tblSaveXML] (
  [ID] AUTOINCREMENT,
  [ObjectType] VARCHAR (255),
  [Notes] VARCHAR (255),
  [AddDate] DATETIME ,
  [UpdateDate] DATETIME ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ID], [ObjectType])
)
