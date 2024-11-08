SELECT
  qryStrings.ID,
  qryStrings.msgid,
  qryStrings.Context,
  qryStrings.Comments,
  tblTranslation.Translation,
  qryStrings.LanguageID AS Lang,
  qryStrings.Reference,
  IIf([msgid] = "", 1, 2) AS SortRank,
  [Context] & "|" & [msgid] AS [Key]
FROM
  qryStrings
  LEFT JOIN tblTranslation ON (
    qryStrings.LanguageID = tblTranslation.Language
  )
  AND (
    qryStrings.ID = tblTranslation.StringID
  )
ORDER BY
  IIf([msgid] = "", 1, 2),
  qryStrings.msgid;
