CREATE TABLE [USysApplicationLog] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SourceObject] VARCHAR (255),
  [Data Macro Instance ID] VARCHAR (255),
  [Error Number] LONG ,
  [Category] VARCHAR (255),
  [Object Type] VARCHAR (255),
  [Description] LONGTEXT ,
  [Context] VARCHAR (255),
  [Created] DATETIME 
)
