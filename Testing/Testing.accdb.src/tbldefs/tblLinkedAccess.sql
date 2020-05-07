CREATE TABLE [tblLinkedAccess] (
  [ID] AUTOINCREMENT,
  [ObjectType] VARCHAR (255),
  [Notes] VARCHAR (255),
   CONSTRAINT [PrimaryKey] PRIMARY KEY (ID, ObjectType)
)
