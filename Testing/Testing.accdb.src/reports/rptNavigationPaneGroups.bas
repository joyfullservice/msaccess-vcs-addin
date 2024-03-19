Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11700
    DatasheetFontHeight =11
    ItemSuffix =7
    RecSrcDt = Begin
        0x2df649898e77e540
    End
    RecordSource ="qryNavigationPaneGroups"
    Caption ="qryNavigationPaneGroups"
    DatasheetFontName ="Franklin Gothic Book"
    FilterOnLoad =0
    FitToPage =1
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
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="ObjectName"
        End
        Begin BreakLevel
            ControlSource ="GroupName"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =960
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    TextFontFamily =0
                    Left =60
                    Top =60
                    Width =4350
                    Height =540
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label4"
                    Caption ="qryNavigationPaneGroups"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =600
                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =0
                    Left =360
                    Top =60
                    Width =3660
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="ObjectName_Label"
                    Caption ="ObjectName"
                    FontName ="Franklin Gothic Book"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =0
                    Left =4380
                    Top =60
                    Width =7260
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="GroupName_Label"
                    Caption ="GroupName"
                    FontName ="Franklin Gothic Book"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4380
                    LayoutCachedTop =60
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =375
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =390
            Name ="GroupHeader0"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =360
                    Width =3660
                    Height =330
                    ColumnWidth =1665
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ObjectName"
                    ControlSource ="ObjectName"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =330
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =4380
                    Width =7260
                    Height =330
                    ColumnWidth =1545
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="GroupName"
                    ControlSource ="GroupName"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =330
                End
            End
        End
        Begin PageFooter
            Height =570
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =5040
                    Height =330
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text5"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =240
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =6600
                    Top =240
                    Width =5040
                    Height =330
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text6"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =6600
                    LayoutCachedTop =240
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =570
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "rptNavigationPaneGroups.cls"
