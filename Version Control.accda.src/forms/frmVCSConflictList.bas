Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ViewsAllowed =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5040
    DatasheetFontHeight =11
    ItemSuffix =31
    Right =19905
    Bottom =12840
    RecSrcDt = Begin
        0x9bf1b7f2f3a6e540
    End
    RecordSource ="tblConflicts"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowFormView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =4117
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1695
                    Top =360
                    Width =2625
                    Height =360
                    ColumnWidth =1815
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    Name ="txtComponent"
                    ControlSource ="Component"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1695
                    LayoutCachedTop =360
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =720
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =360
                            Width =1275
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            Name ="Label3"
                            Caption ="Component"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =720
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1695
                    Top =900
                    Width =2625
                    Height =360
                    ColumnWidth =2595
                    TabIndex =1
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    Name ="txtFileName"
                    ControlSource ="FileName"
                    GroupTable =1
                    BottomPadding =150
                    AggregateType =2

                    LayoutCachedLeft =1695
                    LayoutCachedTop =900
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1260
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =900
                            Width =1275
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            Name ="Label6"
                            Caption ="FileName"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =360
                            LayoutCachedTop =900
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =1260
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1695
                    Top =1440
                    Width =2625
                    Height =360
                    ColumnWidth =1785
                    TabIndex =2
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    Name ="txtObjectDate"
                    ControlSource ="ObjectDate"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1695
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1800
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =1440
                            Width =1275
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            Name ="Label9"
                            Caption ="ObjectDate"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =360
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =1800
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1695
                    Top =1980
                    Width =2625
                    Height =360
                    ColumnWidth =1755
                    TabIndex =3
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    Name ="txtFileDate"
                    ControlSource ="FileDate"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1695
                    LayoutCachedTop =1980
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =2340
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =1980
                            Width =1275
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            Name ="Label12"
                            Caption ="FileDate"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =360
                            LayoutCachedTop =1980
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =2340
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1695
                    Top =3060
                    Width =2625
                    Height =375
                    TabIndex =5
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    Name ="txtDiff"
                    ControlSource ="Diff"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1695
                    LayoutCachedTop =3060
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =3435
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =3060
                            Width =1275
                            Height =375
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            Name ="Label18"
                            Caption ="Diff"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =360
                            LayoutCachedTop =3060
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =3435
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =1695
                    Top =2520
                    Width =2625
                    Height =360
                    TabIndex =4
                    Name ="cboResolution"
                    ControlSource ="Resolution"
                    RowSourceType ="Value List"
                    ColumnWidths ="0"
                    GroupTable =1
                    BottomPadding =150
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22

                    LayoutCachedLeft =1695
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =2880
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =2520
                            Width =1275
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            Name ="Label15"
                            Caption ="Resolution"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =360
                            LayoutCachedTop =2520
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =2880
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Load list of conflict resolutions
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    With cboResolution
        .AddItem eResolveConflict.ercSkip & ";" & "Skip"
        .AddItem eResolveConflict.ercOverwrite & ";" & "Overwrite"
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtDiff_Click
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Launch diff program to review changes
'---------------------------------------------------------------------------------------
'
Private Sub txtDiff_Click()
    MsgBox "Showing diff for " & txtFileName
End Sub
