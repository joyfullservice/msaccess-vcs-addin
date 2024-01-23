CREATE TABLE [tblConflicts] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Component] VARCHAR (255),
  [ItemKey] VARCHAR (255),
  [FileName] VARCHAR (255),
  [ObjectDate] DATETIME,
  [IndexDate] DATETIME,
  [FileDate] DATETIME,
  [Suggestion] LONG,
  [Resolution] LONG,
  [Diff] LONGTEXT
)
