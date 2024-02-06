CREATE TABLE [tblLinkedAccess] (
  [ID] AUTOINCREMENT,
  [ObjectType] VARCHAR (255),
  [Notes] VARCHAR (255),
  [Index&Test] VARCHAR (255),
  [MyAttachment] VARCHAR,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([ID], [ObjectType])
)
