CREATE TABLE [tblSaveXML] (
  [ID] AUTOINCREMENT,
  [ObjectType] VARCHAR (255),
  [Notes] VARCHAR (255),
  [AddDate] DATETIME ,
  [UpdateDate] DATETIME ,
  [NotReq''d] VARCHAR (255),
  [Please""don''t""use] VARCHAR (255),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ID], [ObjectType])
)
