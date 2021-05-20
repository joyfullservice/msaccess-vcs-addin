SELECT tblStrings.ID, [msgid] & "|" & [Context] AS [Key], tblStrings.Context, tblStrings.msgid, tblTranslation.Translation, tblStrings.Comments, tblTranslation.Language
FROM tblStrings LEFT JOIN tblTranslation ON tblStrings.ID = tblTranslation.StringID
WHERE (((tblTranslation.Language)="en_US")) OR (((tblTranslation.Language) Is Null));
