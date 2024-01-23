CREATE TABLE [tblResources] (
  [ResourceName] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Content] VARCHAR,
  [Description] VARCHAR (255)
)
