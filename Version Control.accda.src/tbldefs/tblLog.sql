CREATE TABLE [tblLog] (
  [EntryID] AUTOINCREMENT,
  [EventLogTime] VARCHAR (40),
  [UserID] VARCHAR (40),
  [ComputerName] VARCHAR (255),
  [ErrorNumber] LONG,
  [EventSource] LONGTEXT,
  [EventMessage] LONGTEXT,
  [Printed] BIT,
  [ErrorLevel] LONG
)
