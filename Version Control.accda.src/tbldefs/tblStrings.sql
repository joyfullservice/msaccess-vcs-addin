CREATE TABLE [tblStrings] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [msgid] VARCHAR (255),
  [FullString] LONGTEXT ,
  [Context] VARCHAR (255),
  [Comments] LONGTEXT 
)
