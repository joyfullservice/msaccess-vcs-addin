SELECT tblStrings.ID, [msgid] & "|" & [Context] AS [Key], tblTranslation.Translation, tblStrings.Comments
FROM tblStrings LEFT JOIN tblTranslation ON tblStrings.ID = tblTranslation.StringID
WHERE (((tblTranslation.Language)=GetCurrentLanguage() Or (tblTranslation.Language) Is Null));
