"""Write compact oracle fixtures that cover cases the harvested forms miss."""

from __future__ import annotations

from pathlib import Path

from form_geometry import write_form


HERE = Path(__file__).resolve().parent
FIXTURES = HERE / "fixtures"


def form(body: str) -> str:
    return (
        "Version =20\r\n"
        "VersionRequired =20\r\n"
        "Begin Form\r\n"
        "    Width =6000\r\n"
        "    Caption =\"DPI Layout Probe\"\r\n"
        "    Begin\r\n"
        f"{body}"
        "    End\r\n"
        "End\r\n"
    )


CUSTOM_GAP = form(
    """        Begin FormHeader
            Height =600
            Name ="FormHeader"
            Begin
                Begin Label
                    Left =120
                    Top =120
                    Width =1800
                    Height =360
                    Name ="lblA"
                    GroupTable =1
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =480
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    Left =2100
                    Top =120
                    Width =1800
                    Height =360
                    Name ="lblB"
                    GroupTable =1
                    LayoutCachedLeft =2100
                    LayoutCachedTop =120
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =480
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
"""
)

ZERO_GAP = form(
    """        Begin FormHeader
            Height =600
            Name ="FormHeader"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =1200
                    Height =300
                    Name ="lblA"
                    GroupTable =1
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =360
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    Left =1260
                    Top =60
                    Width =1200
                    Height =300
                    Name ="lblB"
                    GroupTable =1
                    LayoutCachedLeft =1260
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =360
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
"""
)

COLUMN_SPAN = form(
    """        Begin FormHeader
            Height =1200
            Name ="FormHeader"
            Begin
                Begin Label
                    Left =120
                    Top =60
                    Width =1860
                    Height =300
                    Name ="lblSpan"
                    GroupTable =1
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =360
                    ColumnStart =0
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    Left =120
                    Top =420
                    Width =840
                    Height =300
                    Name ="lblLeft"
                    GroupTable =1
                    LayoutCachedLeft =120
                    LayoutCachedTop =420
                    LayoutCachedWidth =960
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    Left =1080
                    Top =420
                    Width =840
                    Height =300
                    Name ="lblRight"
                    GroupTable =1
                    LayoutCachedLeft =1080
                    LayoutCachedTop =420
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
"""
)

TAB_PAGE = form(
    """        Begin Section
            Height =3600
            Name ="Detail"
            Begin
                Begin Tab
                    Width =4800
                    Height =3000
                    Name ="tabMain"
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =3000
                    Begin
                        Begin Page
                            Left =120
                            Top =420
                            Width =4560
                            Height =2460
                            Name ="pgOne"
                            LayoutCachedLeft =120
                            LayoutCachedTop =420
                            LayoutCachedWidth =4680
                            LayoutCachedHeight =2880
                            Begin
                                Begin Label
                                    Left =180
                                    Top =120
                                    Width =1440
                                    Height =300
                                    Name ="lblInner"
                                    GroupTable =1
                                    LayoutCachedLeft =180
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =1620
                                    LayoutCachedHeight =420
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                                Begin TextBox
                                    Left =1740
                                    Top =120
                                    Width =2160
                                    Height =300
                                    Name ="txtInner"
                                    GroupTable =1
                                    LayoutCachedLeft =1740
                                    LayoutCachedTop =120
                                    LayoutCachedWidth =3900
                                    LayoutCachedHeight =420
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =1
                                    GroupTable =1
                                End
                            End
                        End
                    End
                End
            End
        End
"""
)

CROSS_SECTION = form(
    """        Begin FormHeader
            Height =480
            Name ="FormHeader"
            Begin
                Begin Label
                    Left =105
                    Top =60
                    Width =1800
                    Height =360
                    Name ="lblHead"
                    GroupTable =1
                    LayoutCachedLeft =105
                    LayoutCachedTop =60
                    LayoutCachedWidth =1905
                    LayoutCachedHeight =420
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =105
                    Top =30
                    Width =1800
                    Height =345
                    Name ="txtBody"
                    GroupTable =1
                    LayoutCachedLeft =105
                    LayoutCachedTop =30
                    LayoutCachedWidth =1905
                    LayoutCachedHeight =375
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
"""
)


def main() -> None:
    FIXTURES.mkdir(parents=True, exist_ok=True)
    mapping = {
        "CustomGap.form": CUSTOM_GAP,
        "ZeroGap.form": ZERO_GAP,
        "ColumnSpan.form": COLUMN_SPAN,
        "TabPage.form": TAB_PAGE,
        "CrossSection.form": CROSS_SECTION,
    }
    for name, content in mapping.items():
        write_form(FIXTURES / name, content)
        print(name)


if __name__ == "__main__":
    main()
