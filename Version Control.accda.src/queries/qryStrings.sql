SELECT tblStrings.ID, [Context] & "|" & [msgid] AS [Key], tblStrings.msgid, tblStrings.Context, tblStrings.Comments, tblTranslation.Translation, tblTranslation.Language
FROM tblStrings LEFT JOIN tblTranslation ON tblStrings.ID = tblTranslation.StringID;
