SELECT
  [Forms]![frmColors]![Text18] AS FormControl,
  IIf(
    [Forms]![frmVCSInstall]![chkUseRibbon],
    Eval("True"),
    False
  ) AS TestExpression;
