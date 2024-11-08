SELECT
  tblLanguages.ID AS LanguageID,
  tblStrings.ID,
  tblStrings.msgid,
  tblStrings.Context,
  tblStrings.Reference,
  tblStrings.Comments
FROM
  tblLanguages,
  tblStrings;
