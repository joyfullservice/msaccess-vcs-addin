Attribute VB_Name = "modCommandBarNames"
'---------------------------------------------------------------------------------------
' Module    : modCommandBarNames
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Friendly reference names for built-in CommandBar control Ids (and FaceIds),
'           : plus runtime addability classification for replica export/import.
'           : Used to populate read-only ControlIdName / FaceIdName fields in exported
'           : menu JSON. Names are hard-coded for deterministic export; use
'           : DumpCommandBarControlNames and DumpNonAddableControls to grow the tables.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Components")

Private Const ModuleName As String = "modCommandBarNames"

' Session cache of (Type, Id) -> addable Boolean, keyed "Type|Id". Addability is stable
' for the life of an Access session, so probe results are reused across controls, command
' bars, change-scans, and exports. Reset only matters for tests (ResetAddableCache).
Private m_dAddableCache As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : ControlIdToName
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Return a friendly English name for a built-in CommandBar control Id or FaceId.
'           : Returns vbNullString when the Id is not in the table. Mirrors the pattern
'           : of LanguageIdToString in clsTranslation, but without caching.
'---------------------------------------------------------------------------------------
'
Public Function ControlIdToName(lngId As Long) As String

    Select Case lngId
        Case 2:  ControlIdToName = "Spelling"
        Case 3:  ControlIdToName = "Save"
        Case 4:  ControlIdToName = "Print"
        Case 5:  ControlIdToName = "One Page"
        Case 7:  ControlIdToName = "Zoom 100%"
        Case 19:  ControlIdToName = "Copy"
        Case 21:  ControlIdToName = "Cut"
        Case 22:  ControlIdToName = "Paste"
        Case 25:  ControlIdToName = "Zoom"
        Case 29:  ControlIdToName = "Undelete"
        Case 51:  ControlIdToName = "Breakpoint"
        Case 106:  ControlIdToName = "Close"
        Case 108:  ControlIdToName = "Format Painter"
        Case 109:  ControlIdToName = "Print Preview"
        Case 113:  ControlIdToName = "Bold"
        Case 114:  ControlIdToName = "Italic"
        Case 115:  ControlIdToName = "Underline"
        Case 120:  ControlIdToName = "Align Left"
        Case 121:  ControlIdToName = "Align Right"
        Case 122:  ControlIdToName = "Center"
        Case 127:  ControlIdToName = "Page Numbers"
        Case 128:  ControlIdToName = "Undo"
        Case 129:  ControlIdToName = "Redo"
        Case 130:  ControlIdToName = "Line"
        Case 131:  ControlIdToName = "Rectangle"
        Case 141:  ControlIdToName = "Find"
        Case 154:  ControlIdToName = "First"
        Case 155:  ControlIdToName = "Previous"
        Case 156:  ControlIdToName = "Next"
        Case 157:  ControlIdToName = "Last"
        Case 164:  ControlIdToName = "Group"
        Case 165:  ControlIdToName = "Ungroup"
        Case 166:  ControlIdToName = "Bring to Front"
        Case 167:  ControlIdToName = "Send to Back"
        Case 177:  ControlIdToName = "Multiple Pages"
        Case 182:  ControlIdToName = "Select"
        Case 210:  ControlIdToName = "Sort Ascending"
        Case 211:  ControlIdToName = "Sort Descending"
        Case 212:  ControlIdToName = "View"
        Case 219:  ControlIdToName = "Text Box"
        Case 220:  ControlIdToName = "Check Box"
        Case 221:  ControlIdToName = "Combo Box"
        Case 222:  ControlIdToName = "Properties"
        Case 247:  ControlIdToName = "Page Setup"
        Case 253:  ControlIdToName = "Set Label Font"
        Case 282:  ControlIdToName = "Button"
        Case 293:  ControlIdToName = "Delete Item/Transaction"
        Case 294:  ControlIdToName = "Delete Columns"
        Case 296:  ControlIdToName = "Rows"
        Case 297:  ControlIdToName = "Columns"
        Case 302:  ControlIdToName = "Split"
        Case 313:  ControlIdToName = "Replace"
        Case 446:  ControlIdToName = "Option Button"
        Case 448:  ControlIdToName = "List Box"
        Case 467:  ControlIdToName = "Option Group"
        Case 469:  ControlIdToName = "Tab Order"
        Case 473:  ControlIdToName = "Object Browser"
        Case 476:  ControlIdToName = "Label"
        Case 478:  ControlIdToName = "Delete"
        Case 485:  ControlIdToName = "Grid"
        Case 490:  ControlIdToName = "Line/Border Width"
        Case 491:  ControlIdToName = "Line/Border Width"
        Case 492:  ControlIdToName = "Line/Border Width"
        Case 493:  ControlIdToName = "Line/Border Width"
        Case 494:  ControlIdToName = "Line/Border Width"
        Case 495:  ControlIdToName = "Line/Border Width"
        Case 496:  ControlIdToName = "Line/Border Width"
        Case 497:  ControlIdToName = "Apply Filter/Sort"
        Case 498:  ControlIdToName = "View Datasheet"
        Case 499:  ControlIdToName = "Advanced Filter/Sort"
        Case 500:  ControlIdToName = "Subform/Subreport"
        Case 501:  ControlIdToName = "Add Existing Fields"
        Case 502:  ControlIdToName = "Form View"
        Case 504:  ControlIdToName = "Sorting and Grouping"
        Case 505:  ControlIdToName = "Primary Key"
        Case 507:  ControlIdToName = "Bound Object Frame"
        Case 508:  ControlIdToName = "Unbound Object Frame"
        Case 509:  ControlIdToName = "Page Break"
        Case 512:  ControlIdToName = "Form Design"
        Case 513:  ControlIdToName = "Client Query Wizard"
        Case 514:  ControlIdToName = "Report Design"
        Case 515:  ControlIdToName = "Reset"
        Case 520:  ControlIdToName = "Toggle Button"
        Case 522:  ControlIdToName = "Options"
        Case 523:  ControlIdToName = "Relationships"
        Case 524:  ControlIdToName = "Import"
        Case 527:  ControlIdToName = "Indexes"
        Case 528:  ControlIdToName = "SQL View"
        Case 529:  ControlIdToName = "Table Names"
        Case 530:  ControlIdToName = "Show Table"
        Case 531:  ControlIdToName = "Select Query"
        Case 532:  ControlIdToName = "Crosstab Query"
        Case 533:  ControlIdToName = "Make Table Query"
        Case 534:  ControlIdToName = "Update Query"
        Case 535:  ControlIdToName = "Append Query"
        Case 536:  ControlIdToName = "Delete Query"
        Case 537:  ControlIdToName = "Parameters"
        Case 538:  ControlIdToName = "Save Record"
        Case 539:  ControlIdToName = "Add"
        Case 541:  ControlIdToName = "Row Height"
        Case 542:  ControlIdToName = "Field Width"
        Case 544:  ControlIdToName = "Freeze Fields"
        Case 546:  ControlIdToName = "Object"
        Case 548:  ControlIdToName = "Toolbox"
        Case 551:  ControlIdToName = "Size to Fit"
        Case 556:  ControlIdToName = "Page Header/Footer"
        Case 557:  ControlIdToName = "Form Header/Footer"
        Case 558:  ControlIdToName = "First 10 Records Preview"
        Case 559:  ControlIdToName = "Procedure"
        Case 561:  ControlIdToName = "Special Effect Raised"
        Case 562:  ControlIdToName = "Special Effect Sunken"
        Case 566:  ControlIdToName = "Analyze with Microsoft Excel"
        Case 567:  ControlIdToName = "Publish with Microsoft Word"
        Case 570:  ControlIdToName = "Find Next"
        Case 571:  ControlIdToName = "Duplicate"
        Case 572:  ControlIdToName = "User and Group Permissions"
        Case 574:  ControlIdToName = "Form Header/Footer"
        Case 577:  ControlIdToName = "Navigation Pane"
        Case 580:  ControlIdToName = "Special Effect Flat"
        Case 581:  ControlIdToName = "AutoForm"
        Case 582:  ControlIdToName = "AutoReport"
        Case 583:  ControlIdToName = "Table Design"
        Case 601:  ControlIdToName = "Toggle Filter"
        Case 605:  ControlIdToName = "Remove Filter/Sort"
        Case 613:  ControlIdToName = "Thin"
        Case 614:  ControlIdToName = "Dashed"
        Case 615:  ControlIdToName = "Line/Border Style"
        Case 616:  ControlIdToName = "Dotted"
        Case 617:  ControlIdToName = "Line/Border Style"
        Case 618:  ControlIdToName = "Dash Dot"
        Case 619:  ControlIdToName = "Dash Dot Dot"
        Case 621:  ControlIdToName = "Module"
        Case 622:  ControlIdToName = "Show Direct Relationships"
        Case 623:  ControlIdToName = "Show All Relationships"
        Case 625:  ControlIdToName = "Control Wizards"
        Case 626:  ControlIdToName = "Word Merge"
        Case 628:  ControlIdToName = "Filter By Form"
        Case 629:  ControlIdToName = "Macro"
        Case 630:  ControlIdToName = "AutoFormat"
        Case 631:  ControlIdToName = "Large Icons"
        Case 632:  ControlIdToName = "Small Icons"
        Case 633:  ControlIdToName = "List"
        Case 634:  ControlIdToName = "Details"
        Case 635:  ControlIdToName = "Gridlines Both"
        Case 636:  ControlIdToName = "Gridlines Vertical"
        Case 637:  ControlIdToName = "Gridlines Horizontal"
        Case 638:  ControlIdToName = "Gridlines None"
        Case 639:  ControlIdToName = "Two Pages"
        Case 640:  ControlIdToName = "Filter By Selection"
        Case 642:  ControlIdToName = "More Controls"
        Case 644:  ControlIdToName = "Delete Record"
        Case 645:  ControlIdToName = "Build"
        Case 653:  ControlIdToName = "By Modified"
        Case 654:  ControlIdToName = "By Type"
        Case 655:  ControlIdToName = "By Name"
        Case 656:  ControlIdToName = "By Created"
        Case 658:  ControlIdToName = "Documenter"
        Case 659:  ControlIdToName = "Performance"
        Case 660:  ControlIdToName = "Table"
        Case 664:  ControlIdToName = "Left"
        Case 665:  ControlIdToName = "Right"
        Case 666:  ControlIdToName = "Top"
        Case 667:  ControlIdToName = "Bottom"
        Case 705:  ControlIdToName = "Collapse All"
        Case 706:  ControlIdToName = "Expand All"
        Case 716:  ControlIdToName = "Insert Field"
        Case 717:  ControlIdToName = "Insert Field"
        Case 725:  ControlIdToName = "Collapse Parent"
        Case 747:  ControlIdToName = "Close All"
        Case 748:  ControlIdToName = "Save As"
        Case 752:  ControlIdToName = "Exit"
        Case 755:  ControlIdToName = "Paste Special"
        Case 756:  ControlIdToName = "Select All"
        Case 761:  ControlIdToName = "Data Selector"
        Case 768:  ControlIdToName = "Date and Time"
        Case 793:  ControlIdToName = "AutoCorrect Options"
        Case 797:  ControlIdToName = "Customize"
        Case 809:  ControlIdToName = "Office Clipboard"
        Case 822:  ControlIdToName = "Hide"
        Case 830:  ControlIdToName = "Window Name Goes Here"
        Case 838:  ControlIdToName = "Minimize"
        Case 839:  ControlIdToName = "Restore"
        Case 840:  ControlIdToName = "Close"
        Case 841:  ControlIdToName = "Move"
        Case 842:  ControlIdToName = "Size"
        Case 843:  ControlIdToName = "Maximize"
        Case 855:  ControlIdToName = "Datasheet"
        Case 866:  ControlIdToName = "Unhide"
        Case 925:  ControlIdToName = "Zoom"
        Case 927:  ControlIdToName = "About Microsoft Access"
        Case 931:  ControlIdToName = "Pictures"
        Case 939:  ControlIdToName = "Procedure Definition"
        Case 940:  ControlIdToName = "Edit Watch"
        Case 942:  ControlIdToName = "References"
        Case 959:  ControlIdToName = "More Windows"
        Case 964:  ControlIdToName = "Run"
        Case 965:  ControlIdToName = "Image"
        Case 966:  ControlIdToName = "Ruler"
        Case 967:  ControlIdToName = "Grid"
        Case 983:  ControlIdToName = "Contents and Index"
        Case 984:  ControlIdToName = "Help"
        Case 1015:  ControlIdToName = "Open Hyperlink"
        Case 1016:  ControlIdToName = "Start Page"
        Case 1017:  ControlIdToName = "Back"
        Case 1018:  ControlIdToName = "Forward"
        Case 1019:  ControlIdToName = "Stop"
        Case 1020:  ControlIdToName = "Refresh"
        Case 1021:  ControlIdToName = "Open Favorites"
        Case 1022:  ControlIdToName = "Add to Favorites"
        Case 1023:  ControlIdToName = "Show Only Web Toolbar"
        Case 1081:  ControlIdToName = "Design View"
        Case 1083:  ControlIdToName = "Delete Watch"
        Case 1435:  ControlIdToName = "Arrange Icons"
        Case 1457:  ControlIdToName = "Recently Used Colors"
        Case 1567:  ControlIdToName = "Close Master View"
        Case 1568:  ControlIdToName = "Create Menu from Macro"
        Case 1569:  ControlIdToName = "Create Toolbar from Macro"
        Case 1570:  ControlIdToName = "Create Shortcut Menu from Macro"
        Case 1574:  ControlIdToName = "Open in New Window"
        Case 1575:  ControlIdToName = "Copy Hyperlink"
        Case 1576:  ControlIdToName = "Hyperlink"
        Case 1577:  ControlIdToName = "Edit Hyperlink"
        Case 1649:  ControlIdToName = "Decrease"
        Case 1650:  ControlIdToName = "Increase"
        Case 1652:  ControlIdToName = "Decrease"
        Case 1653:  ControlIdToName = "Increase"
        Case 1695:  ControlIdToName = "Visual Basic Editor"
        Case 1728:  ControlIdToName = "Font"
        Case 1731:  ControlIdToName = "Font Size"
        Case 1733:  ControlIdToName = "Zoom"
        Case 1734:  ControlIdToName = "Object"
        Case 1736:  ControlIdToName = "Go To Field"
        Case 1740:  ControlIdToName = "Address"
        Case 1766:  ControlIdToName = "Query Type"
        Case 1767:  ControlIdToName = "Line/Border Width"
        Case 1769:  ControlIdToName = "Special Effect"
        Case 1772:  ControlIdToName = "Top Values"
        Case 1773:  ControlIdToName = "Tables"
        Case 1774:  ControlIdToName = "Queries"
        Case 1775:  ControlIdToName = "Forms"
        Case 1776:  ControlIdToName = "Reports"
        Case 1777:  ControlIdToName = "Macros"
        Case 1778:  ControlIdToName = "Modules"
        Case 1780:  ControlIdToName = "Select Form"
        Case 1781:  ControlIdToName = "Select Report"
        Case 1782:  ControlIdToName = "Select Record"
        Case 1784:  ControlIdToName = "Select All Records"
        Case 1785:  ControlIdToName = "Test Validation Rules"
        Case 1786:  ControlIdToName = "Clear Layout"
        Case 1787:  ControlIdToName = "Rename"
        Case 1788:  ControlIdToName = "Create Shortcut"
        Case 1789:  ControlIdToName = "Save Layout"
        Case 1790:  ControlIdToName = "Save As Query"
        Case 1791:  ControlIdToName = "Load from Query"
        Case 1792:  ControlIdToName = "Save as Text"
        Case 1793:  ControlIdToName = "Rename Field"
        Case 1794:  ControlIdToName = "Unfreeze All Fields"
        Case 1795:  ControlIdToName = "Link Tables"
        Case 1796:  ControlIdToName = "Modify Lookups"
        Case 1802:  ControlIdToName = "Remove Table"
        Case 1803:  ControlIdToName = "Data Entry"
        Case 1804:  ControlIdToName = "Refresh"
        Case 1805:  ControlIdToName = "Edit Relationship"
        Case 1806:  ControlIdToName = "Hide Table"
        Case 1811:  ControlIdToName = "Run to Cursor"
        Case 1812:  ControlIdToName = "Set Next Statement"
        Case 1813:  ControlIdToName = "Show Next Statement"
        Case 1814:  ControlIdToName = "Encrypt with Password"
        Case 1815:  ControlIdToName = "User-Level Security Wizard"
        Case 1816:  ControlIdToName = "User and Group Accounts"
        Case 1817:  ControlIdToName = "Special Effect Etched"
        Case 1818:  ControlIdToName = "Special Effect Shadowed"
        Case 1819:  ControlIdToName = "Special Effect Chiseled"
        Case 1820:  ControlIdToName = "Add Watch"
        Case 1821:  ControlIdToName = "Startup"
        Case 1822:  ControlIdToName = "Last Position"
        Case 1823:  ControlIdToName = "Datasheet"
        Case 1824:  ControlIdToName = "Join Properties"
        Case 1825:  ControlIdToName = "Line Up Icons"
        Case 1826:  ControlIdToName = "Cascade"
        Case 1828:  ControlIdToName = "Size to Fit Form"
        Case 1829:  ControlIdToName = "200%"
        Case 1830:  ControlIdToName = "150%"
        Case 1831:  ControlIdToName = "75%"
        Case 1832:  ControlIdToName = "50%"
        Case 1833:  ControlIdToName = "25%"
        Case 1834:  ControlIdToName = "10%"
        Case 1835:  ControlIdToName = "Open Table"
        Case 1836:  ControlIdToName = "Open Query"
        Case 1837:  ControlIdToName = "Open Form"
        Case 1838:  ControlIdToName = "Open Report"
        Case 1839:  ControlIdToName = "About"
        Case 1922:  ControlIdToName = "Search the Web"
        Case 1955:  ControlIdToName = "Hide Fields"
        Case 1957:  ControlIdToName = "Chart"
        Case 1965:  ControlIdToName = "<verb>"
        Case 1967:  ControlIdToName = "Convert"
        Case 2037:  ControlIdToName = "Properties"
        Case 2043:  ControlIdToName = "Query Design"
        Case 2057:  ControlIdToName = "Set Start Page"
        Case 2058:  ControlIdToName = "Set Search Page"
        Case 2069:  ControlIdToName = "Auto Arrange"
        Case 2070:  ControlIdToName = "To Access 2002 - 2003 File Format"
        Case 2071:  ControlIdToName = "Compact and Repair Database"
        Case 2073:  ControlIdToName = "Make ACCDE"
        Case 2074:  ControlIdToName = "Encode/Decode Database"
        Case 2075:  ControlIdToName = "Paste Append"
        Case 2076:  ControlIdToName = "Delete Tab"
        Case 2077:  ControlIdToName = "To Tallest"
        Case 2078:  ControlIdToName = "To Shortest"
        Case 2079:  ControlIdToName = "To Widest"
        Case 2080:  ControlIdToName = "To Narrowest"
        Case 2081:  ControlIdToName = "Union"
        Case 2082:  ControlIdToName = "Data Definition"
        Case 2083:  ControlIdToName = "Pass-Through"
        Case 2084:  ControlIdToName = "Transparent"
        Case 2085:  ControlIdToName = "Transparent"
        Case 2086:  ControlIdToName = "Build Event"
        Case 2087:  ControlIdToName = "Clear Grid"
        Case 2161:  ControlIdToName = "Insert Rows"
        Case 2165:  ControlIdToName = "Delete Rows"
        Case 2188:  ControlIdToName = "Mail Recipient (as Attachment)"
        Case 2521:  ControlIdToName = "Print"
        Case 2525:  ControlIdToName = "Bookmark"
        Case 2526:  ControlIdToName = "Next Bookmark"
        Case 2527:  ControlIdToName = "Previous Bookmark"
        Case 2528:  ControlIdToName = "Clear All Bookmarks"
        Case 2529:  ControlIdToName = "List Properties/Methods"
        Case 2530:  ControlIdToName = "List Constants"
        Case 2531:  ControlIdToName = "Quick Info"
        Case 2532:  ControlIdToName = "Parameter Info"
        Case 2533:  ControlIdToName = "Complete Word"
        Case 2561:  ControlIdToName = "Tile Vertically"
        Case 2562:  ControlIdToName = "Tile Horizontally"
        Case 2579:  ControlIdToName = "Class Module"
        Case 2585:  ControlIdToName = "Tab Control"
        Case 2591:  ControlIdToName = "Insert Page"
        Case 2592:  ControlIdToName = "Delete Page"
        Case 2598:  ControlIdToName = "Office Links"
        Case 2599:  ControlIdToName = "New Object"
        Case 2623:  ControlIdToName = "Add-In Manager"
        Case 2646:  ControlIdToName = "Add Objects to SourceSafe"
        Case 2647:  ControlIdToName = "Get Latest Version"
        Case 2648:  ControlIdToName = "Check Out"
        Case 2649:  ControlIdToName = "Check In"
        Case 2650:  ControlIdToName = "Undo Check Out"
        Case 2651:  ControlIdToName = "Share Objects"
        Case 2652:  ControlIdToName = "Show Differences"
        Case 2653:  ControlIdToName = "Show History"
        Case 2654:  ControlIdToName = "Run SourceSafe"
        Case 2655:  ControlIdToName = "SourceSafe Properties"
        Case 2656:  ControlIdToName = "Create Database from SourceSafe Project"
        Case 2657:  ControlIdToName = "Add Database to SourceSafe"
        Case 2658:  ControlIdToName = "Options"
        Case 2659:  ControlIdToName = "Refresh Status"
        Case 2660:  ControlIdToName = "Publish to the Web"
        Case 2664:  ControlIdToName = "Open Hyperlink"
        Case 2666:  ControlIdToName = "Analyze"
        Case 2669:  ControlIdToName = "Get Latest Version"
        Case 2670:  ControlIdToName = "Check Out"
        Case 2671:  ControlIdToName = "Undo Check Out"
        Case 2689:  ControlIdToName = "Show Hidden Members"
        Case 2690:  ControlIdToName = "View Definition"
        Case 2763:  ControlIdToName = "Page Order"
        Case 2764:  ControlIdToName = "Unhide Fields"
        Case 2771:  ControlIdToName = "Properties"
        Case 2787:  ControlIdToName = "Paste as Hyperlink"
        Case 2931:  ControlIdToName = "Set Control Defaults"
        Case 2932:  ControlIdToName = "ActiveX Control"
        Case 2936:  ControlIdToName = "New"
        Case 2937:  ControlIdToName = "Open"
        Case 2938:  ControlIdToName = "Export"
        Case 2939:  ControlIdToName = "Database Properties"
        Case 2940:  ControlIdToName = "Delete Field"
        Case 2941:  ControlIdToName = "Select All"
        Case 2942:  ControlIdToName = "OLE/DDE Links"
        Case 2943:  ControlIdToName = "Totals"
        Case 2948:  ControlIdToName = "Convert Macros to Visual Basic"
        Case 2952:  ControlIdToName = "Design View"
        Case 2953:  ControlIdToName = "Find Whole Word Only"
        Case 2993:  ControlIdToName = "Snap to Grid"
        Case 2995:  ControlIdToName = "Insert Object"
        Case 2997:  ControlIdToName = "Go/Continue"
        Case 2998:  ControlIdToName = "End"
        Case 2999:  ControlIdToName = "Single Step"
        Case 3000:  ControlIdToName = "ActiveX Controls"
        Case 3017:  ControlIdToName = "Filter Excluding Selection"
        Case 3058:  ControlIdToName = "Conditional Formatting"
        Case 3076:  ControlIdToName = "Font/Fore Color"
        Case 3077:  ControlIdToName = "Fill/Back Color"
        Case 3078:  ControlIdToName = "Line/Border Color"
        Case 3079:  ControlIdToName = "Font/Fore Color"
        Case 3080:  ControlIdToName = "Font/Fore Color"
        Case 3081:  ControlIdToName = "Fill/Back Color"
        Case 3082:  ControlIdToName = "Fill/Back Color"
        Case 3083:  ControlIdToName = "Line/Border Color"
        Case 3084:  ControlIdToName = "Line/Border Color"
        Case 3109:  ControlIdToName = "Help"
        Case 3110:  ControlIdToName = "Check Box"
        Case 3111:  ControlIdToName = "Combo Box"
        Case 3112:  ControlIdToName = "Command Button"
        Case 3113:  ControlIdToName = "Image"
        Case 3114:  ControlIdToName = "Label"
        Case 3115:  ControlIdToName = "List Box"
        Case 3116:  ControlIdToName = "Option Button"
        Case 3117:  ControlIdToName = "Text Box"
        Case 3118:  ControlIdToName = "Toggle Button"
        Case 3120:  ControlIdToName = "To Grid"
        Case 3121:  ControlIdToName = "To Fit"
        Case 3137:  ControlIdToName = "Make Equal"
        Case 3138:  ControlIdToName = "Make Equal"
        Case 3142:  ControlIdToName = "To Grid"
        Case 3156:  ControlIdToName = "File"
        Case 3164:  ControlIdToName = "Group Members"
        Case 3168:  ControlIdToName = "Lookup Field"
        Case 3169:  ControlIdToName = "Hyperlink Column"
        Case 3170:  ControlIdToName = "Code"
        Case 3227:  ControlIdToName = "Edit Hyperlink"
        Case 3492:  ControlIdToName = "Hangul Hanja Conversion"
        Case 3626:  ControlIdToName = "Remove Hyperlink"
        Case 3627:  ControlIdToName = "Security"
        Case 3655:  ControlIdToName = "Web Page Preview"
        Case 3720:  ControlIdToName = "Reconvert"
        Case 3727:  ControlIdToName = "Meet Now"
        Case 3738:  ControlIdToName = "Mail Recipient"
        Case 3746:  ControlIdToName = "Envelope"
        Case 3774:  ControlIdToName = "Detect and Repair"
        Case 3775:  ControlIdToName = "Office.com"
        Case 3812:  ControlIdToName = "Refresh"
        Case 3828:  ControlIdToName = "Page"
        Case 3832:  ControlIdToName = "Open"
        Case 3840:  ControlIdToName = "Print Relationships"
        Case 3843:  ControlIdToName = "Pages"
        Case 3850:  ControlIdToName = "Connection"
        Case 3851:  ControlIdToName = "Remove"
        Case 3852:  ControlIdToName = "Subdatasheet"
        Case 3857:  ControlIdToName = "New Label"
        Case 3858:  ControlIdToName = "Add Related Tables"
        Case 3859:  ControlIdToName = "Show Relationship Labels"
        Case 3860:  ControlIdToName = "Modify Custom View"
        Case 3861:  ControlIdToName = "View Page Breaks"
        Case 3862:  ControlIdToName = "Recalculate Page Breaks"
        Case 3864:  ControlIdToName = "Arrange Selection"
        Case 3865:  ControlIdToName = "Arrange Tables"
        Case 3866:  ControlIdToName = "New Table"
        Case 3872:  ControlIdToName = "Column Properties"
        Case 3873:  ControlIdToName = "Column Names"
        Case 3874:  ControlIdToName = "Keys"
        Case 3875:  ControlIdToName = "Name Only"
        Case 3876:  ControlIdToName = "Custom View"
        Case 3880:  ControlIdToName = "Delete Table from Database"
        Case 3881:  ControlIdToName = "Hide Table"
        Case 3882:  ControlIdToName = "Autosize Selected Tables"
        Case 3883:  ControlIdToName = "Zoom Selection"
        Case 3884:  ControlIdToName = "Delete Relationship from Database"
        Case 3885:  ControlIdToName = "Open Data Page"
        Case 3886:  ControlIdToName = "Open Server View"
        Case 3888:  ControlIdToName = "Run Stored Procedure"
        Case 3938:  ControlIdToName = "Apply Server Filter"
        Case 3965:  ControlIdToName = "Revert"
        Case 3971:  ControlIdToName = "Server Filter by Form"
        Case 3975:  ControlIdToName = "Diagrams"
        Case 3978:  ControlIdToName = "Verify SQL Syntax"
        Case 3980:  ControlIdToName = "Table Modes"
        Case 3981:  ControlIdToName = "Zoom Modes Dropdown"
        Case 3982:  ControlIdToName = "Diagram"
        Case 3984:  ControlIdToName = "View Design"
        Case 3985:  ControlIdToName = "Stored Procedure Design"
        Case 3986:  ControlIdToName = "Expand All"
        Case 3987:  ControlIdToName = "Collapse All"
        Case 3996:  ControlIdToName = "Linked Table Manager"
        Case 4006:  ControlIdToName = "Select All Columns"
        Case 4007:  ControlIdToName = "Select All Rows from <tableA>"
        Case 4008:  ControlIdToName = "Select All Rows from <tableB>"
        Case 4009:  ControlIdToName = "Remove"
        Case 4010:  ControlIdToName = "Group By"
        Case 4011:  ControlIdToName = "Hide Pane"
        Case 4012:  ControlIdToName = "Diagram"
        Case 4013:  ControlIdToName = "Grid"
        Case 4014:  ControlIdToName = "SQL"
        Case 4015:  ControlIdToName = "Add to Output"
        Case 4016:  ControlIdToName = "Sort Ascending"
        Case 4017:  ControlIdToName = "Sort Descending"
        Case 4018:  ControlIdToName = "Database Splitter"
        Case 4033:  ControlIdToName = "Triggers"
        Case 4078:  ControlIdToName = "More Groups"
        Case 4083:  ControlIdToName = "Custom Group List"
        Case 4133:  ControlIdToName = "Back Up SQL Database"
        Case 4134:  ControlIdToName = "Restore SQL Database"
        Case 4135:  ControlIdToName = "Make ADE File"
        Case 4160:  ControlIdToName = "Switchboard Manager"
        Case 5745:  ControlIdToName = "Run Macro"
        Case 5746:  ControlIdToName = "Task Pane"
        Case 5748:  ControlIdToName = "Server Properties"
        Case 5800:  ControlIdToName = "Subform in New Window"
        Case 5801:  ControlIdToName = "Subreport in New Window"
        Case 5802:  ControlIdToName = "Setup"
        Case 5803:  ControlIdToName = "Average"
        Case 5804:  ControlIdToName = "Count Values"
        Case 5806:  ControlIdToName = "Max"
        Case 5807:  ControlIdToName = "Min"
        Case 5814:  ControlIdToName = "PivotTable View"
        Case 5815:  ControlIdToName = "PivotChart View"
        Case 5865:  ControlIdToName = "Edit SQL"
        Case 5884:  ControlIdToName = "AutoFilter"
        Case 5886:  ControlIdToName = "Sum"
        Case 5887:  ControlIdToName = "Count"
        Case 5888:  ControlIdToName = "Min"
        Case 5889:  ControlIdToName = "Max"
        Case 5890:  ControlIdToName = "Subtotal"
        Case 5891:  ControlIdToName = "Refresh"
        Case 5896:  ControlIdToName = "Expand"
        Case 5897:  ControlIdToName = "Export to Excel"
        Case 5898:  ControlIdToName = "Expand Indicators"
        Case 5899:  ControlIdToName = "Drop Areas"
        Case 5905:  ControlIdToName = "File Search"
        Case 5933:  ControlIdToName = "Activate Product"
        Case 5950:  ControlIdToName = "Sign out"
        Case 5959:  ControlIdToName = "Workgroup Administrator"
        Case 5963:  ControlIdToName = "Size Diagram to Window"
        Case 5964:  ControlIdToName = "Add Table"
        Case 6066:  ControlIdToName = "Copy Database File"
        Case 6067:  ControlIdToName = "Transfer Database"
        Case 6427:  ControlIdToName = "1"
        Case 6428:  ControlIdToName = "2"
        Case 6429:  ControlIdToName = "5"
        Case 6430:  ControlIdToName = "10"
        Case 6431:  ControlIdToName = "25"
        Case 6432:  ControlIdToName = "1%"
        Case 6433:  ControlIdToName = "2%"
        Case 6434:  ControlIdToName = "5%"
        Case 6435:  ControlIdToName = "10%"
        Case 6436:  ControlIdToName = "25%"
        Case 6437:  ControlIdToName = "Other"
        Case 6438:  ControlIdToName = "Show All"
        Case 6439:  ControlIdToName = "Average"
        Case 6440:  ControlIdToName = "Standard Deviation"
        Case 6441:  ControlIdToName = "Standard Deviation Population"
        Case 6442:  ControlIdToName = "Variance"
        Case 6443:  ControlIdToName = "Variance Population"
        Case 6444:  ControlIdToName = "Create Calculated Total"
        Case 6446:  ControlIdToName = "Create Calculated Detail Field"
        Case 6448:  ControlIdToName = "AutoCalc"
        Case 6450:  ControlIdToName = "Change Chart Type"
        Case 6451:  ControlIdToName = "Multiple Plots"
        Case 6452:  ControlIdToName = "Multiple Plots Unified Scale"
        Case 6454:  ControlIdToName = "Switch Row/Column"
        Case 6455:  ControlIdToName = "Drill Into"
        Case 6463:  ControlIdToName = "500%"
        Case 6464:  ControlIdToName = "Maximum 1000%"
        Case 6699:  ControlIdToName = "Calculated Totals and Fields"
        Case 6700:  ControlIdToName = "Show Top/Bottom Items"
        Case 6701:  ControlIdToName = "Collapse"
        Case 6703:  ControlIdToName = "Show Details"
        Case 6704:  ControlIdToName = "Hide Details"
        Case 6705:  ControlIdToName = "Undo"
        Case 6706:  ControlIdToName = "Delete"
        Case 6707:  ControlIdToName = "Show Legend"
        Case 6715:  ControlIdToName = "Form"
        Case 6831:  ControlIdToName = "Show As"
        Case 6832:  ControlIdToName = "Normal"
        Case 6833:  ControlIdToName = "Percent of Row Total"
        Case 6834:  ControlIdToName = "Percent of Column Total"
        Case 6835:  ControlIdToName = "Percent of Parent Row Item"
        Case 6836:  ControlIdToName = "Percent of Parent Column Item"
        Case 6837:  ControlIdToName = "Percent of Grand Total"
        Case 6951:  ControlIdToName = "Select Query"
        Case 6952:  ControlIdToName = "Make-Table Query"
        Case 6953:  ControlIdToName = "Update Query"
        Case 6954:  ControlIdToName = "Append Query"
        Case 6955:  ControlIdToName = "Append Values Query"
        Case 6956:  ControlIdToName = "Delete Query"
        Case 6958:  ControlIdToName = "Maximum Records"
        Case 7043:  ControlIdToName = "Clear Custom Ordering"
        Case 7044:  ControlIdToName = "Filter By Selection"
        Case 7045:  ControlIdToName = "Remove"
        Case 7046:  ControlIdToName = "Group Items"
        Case 7047:  ControlIdToName = "Ungroup Items"
        Case 7048:  ControlIdToName = "Drill Out"
        Case 7049:  ControlIdToName = "Manage Indexes"
        Case 7058:  ControlIdToName = "Remove Filter"
        Case 7059:  ControlIdToName = "Indexes / Keys"
        Case 7060:  ControlIdToName = "Relationships"
        Case 7061:  ControlIdToName = "Constraints"
        Case 7062:  ControlIdToName = "Copy Diagram to Clipboard"
        Case 7096:  ControlIdToName = "Ascending by Total"
        Case 7097:  ControlIdToName = "Descending by Total"
        Case 7490:  ControlIdToName = "Object Dependencies"
        Case 7714:  ControlIdToName = "Privacy Options"
        Case 7903:  ControlIdToName = "Contact Support"
        Case 9340:  ControlIdToName = "Check for Updates"
        Case 9714:  ControlIdToName = "Access Developer Resources"
        Case 10003:  ControlIdToName = "Back Up Database"
        Case 10058:  ControlIdToName = "Next Month"
        Case 10059:  ControlIdToName = "Next Quarter"
        Case 10060:  ControlIdToName = "Next Year"
        Case 10061:  ControlIdToName = "Year To Date"
        Case 10062:  ControlIdToName = "Between"
        Case 10063:  ControlIdToName = "Tomorrow"
        Case 10064:  ControlIdToName = "Today"
        Case 10065:  ControlIdToName = "Yesterday"
        Case 10066:  ControlIdToName = "Last Week"
        Case 10067:  ControlIdToName = "This Week"
        Case 10069:  ControlIdToName = "This Month"
        Case 10070:  ControlIdToName = "Last Month"
        Case 10072:  ControlIdToName = "This Quarter"
        Case 10073:  ControlIdToName = "Last Quarter"
        Case 10074:  ControlIdToName = "This Year"
        Case 10075:  ControlIdToName = "Last Year"
        Case 10077:  ControlIdToName = "Equals"
        Case 10078:  ControlIdToName = "Does Not Equal"
        Case 10079:  ControlIdToName = "Begins With"
        Case 10080:  ControlIdToName = "Contains"
        Case 10081:  ControlIdToName = "Does Not Contain"
        Case 10082:  ControlIdToName = "Less Than or Equal To"
        Case 10083:  ControlIdToName = "Greater Than or Equal To"
        Case 10084:  ControlIdToName = "Quarter 1"
        Case 10085:  ControlIdToName = "Quarter 2"
        Case 10086:  ControlIdToName = "Quarter 3"
        Case 10087:  ControlIdToName = "Quarter 4"
        Case 10088:  ControlIdToName = "Ends With"
        Case 11108:  ControlIdToName = "Sort Descending"
        Case 11111:  ControlIdToName = "Show all groups"
        Case 11112:  ControlIdToName = "Navigation Options"
        Case 11113:  ControlIdToName = "Expand Group"
        Case 11114:  ControlIdToName = "Collapse Group"
        Case 11115:  ControlIdToName = "Show Only This Group"
        Case 11116:  ControlIdToName = "Hide in this Group"
        Case 11117:  ControlIdToName = "Unhide in this Group"
        Case 11118:  ControlIdToName = "Object Type"
        Case 11119:  ControlIdToName = "Tables and Related Views"
        Case 11120:  ControlIdToName = "Created Date"
        Case 11121:  ControlIdToName = "Modified Date"
        Case 11122:  ControlIdToName = "Name"
        Case 11123:  ControlIdToName = "Type"
        Case 11124:  ControlIdToName = "Created Date"
        Case 11125:  ControlIdToName = "Modified Date"
        Case 11126:  ControlIdToName = "List"
        Case 11127:  ControlIdToName = "Icon"
        Case 11128:  ControlIdToName = "Details"
        Case 11253:  ControlIdToName = "Report View"
        Case 11265:  ControlIdToName = "New Group"
        Case 11266:  ControlIdToName = "Rename Shortcut"
        Case 11267:  ControlIdToName = "Remove"
        Case 11268:  ControlIdToName = "Close Database"
        Case 11305:  ControlIdToName = "Design Tasks Pane"
        Case 11666:  ControlIdToName = "Tabular"
        Case 11667:  ControlIdToName = "Stacked"
        Case 11668:  ControlIdToName = "Remove Layout"
        Case 11669:  ControlIdToName = "Select Entire Row"
        Case 11670:  ControlIdToName = "Select Entire Column"
        Case 11673:  ControlIdToName = "Both"
        Case 11674:  ControlIdToName = "Vertical"
        Case 11675:  ControlIdToName = "Horizontal"
        Case 11676:  ControlIdToName = "None"
        Case 11678:  ControlIdToName = "Top"
        Case 11679:  ControlIdToName = "Bottom"
        Case 11711:  ControlIdToName = "Access Database"
        Case 11712:  ControlIdToName = "Excel"
        Case 11713:  ControlIdToName = "Text File"
        Case 11714:  ControlIdToName = "SharePoint List"
        Case 11715:  ControlIdToName = "XML File"
        Case 11716:  ControlIdToName = "ODBC Database"
        Case 11717:  ControlIdToName = "HTML Document"
        Case 11718:  ControlIdToName = "Outlook Folder"
        Case 11723:  ControlIdToName = "Excel"
        Case 11724:  ControlIdToName = "SharePoint List"
        Case 11725:  ControlIdToName = "Word RTF File"
        Case 11726:  ControlIdToName = "Access"
        Case 11727:  ControlIdToName = "Text File"
        Case 11728:  ControlIdToName = "XML File"
        Case 11729:  ControlIdToName = "ODBC Database"
        Case 11731:  ControlIdToName = "HTML Document"
        Case 11732:  ControlIdToName = "dBASE File"
        Case 11751:  ControlIdToName = "Work Online"
        Case 11752:  ControlIdToName = "Synchronize All"
        Case 11753:  ControlIdToName = "Discard Changes and Refresh All"
        Case 12076:  ControlIdToName = "Group On"
        Case 12077:  ControlIdToName = "Sum"
        Case 12329:  ControlIdToName = "Datasheet View"
        Case 12335:  ControlIdToName = "Show All Actions"
        Case 12499:  ControlIdToName = "PDF or XPS"
        Case 12506:  ControlIdToName = "Refresh Link"
        Case 12547:  ControlIdToName = "Recent File Name Goes Here"
        Case 12614:  ControlIdToName = "Past"
        Case 12615:  ControlIdToName = "Future"
        Case 12616:  ControlIdToName = "Paste Formatting"
        Case 12687:  ControlIdToName = "Modify Columns and Settings"
        Case 12688:  ControlIdToName = "Alert Me"
        Case 12689:  ControlIdToName = "Modify Workflow"
        Case 12690:  ControlIdToName = "Change Permissions for this List"
        Case 12691:  ControlIdToName = "Refresh List"
        Case 12696:  ControlIdToName = "Does Not Begin With"
        Case 12697:  ControlIdToName = "Does Not End With"
        Case 12698:  ControlIdToName = "Before"
        Case 12699:  ControlIdToName = "After"
        Case 12898:  ControlIdToName = "Forward"
        Case 12899:  ControlIdToName = "Back"
        Case 12900:  ControlIdToName = "Manage Attachments"
        Case 12950:  ControlIdToName = "Edit List Items"
        Case 12951:  ControlIdToName = "PDF or XPS"
        Case 12952:  ControlIdToName = "Show column history"
        Case 12955:  ControlIdToName = "Access 2007 Database"
        Case 13157:  ControlIdToName = "Layout View"
        Case 13275:  ControlIdToName = "Remove Automatic Sorts"
        Case 13276:  ControlIdToName = "Anchoring"
        Case 13623:  ControlIdToName = "Delete List"
        Case 13624:  ControlIdToName = "Open Default View"
        Case 14186:  ControlIdToName = "Alternate Row Colors"
        Case 14187:  ControlIdToName = "Alternate Fill/Back Color"
        Case 14205:  ControlIdToName = "Alternate Fill/Back Color"
        Case 14398:  ControlIdToName = "Top Left"
        Case 14399:  ControlIdToName = "Stretch Across Top"
        Case 14400:  ControlIdToName = "Top Right"
        Case 14401:  ControlIdToName = "Stretch Down"
        Case 14402:  ControlIdToName = "Stretch Down and Across"
        Case 14403:  ControlIdToName = "Stretch Down and Right"
        Case 14404:  ControlIdToName = "Bottom Left"
        Case 14405:  ControlIdToName = "Stretch Across Bottom"
        Case 14406:  ControlIdToName = "Bottom Right"
        Case 14427:  ControlIdToName = "Search Bar"
        Case 14442:  ControlIdToName = "Field List"
        Case 14467:  ControlIdToName = "Hide"
        Case 14468:  ControlIdToName = "Unhide"
        Case 14782:  ControlIdToName = "Close"
        Case 15056:  ControlIdToName = "Set Caption"
        Case 15113:  ControlIdToName = "None"
        Case 15132:  ControlIdToName = "Standard Deviation"
        Case 15133:  ControlIdToName = "Variance"
        Case 15135:  ControlIdToName = "ADP Custom Groups"
        Case 15169:  ControlIdToName = "Delete"
        Case 15195:  ControlIdToName = "Is Selected"
        Case 15196:  ControlIdToName = "Is Not Selected"
        Case 15681:  ControlIdToName = "Fit to Window"
        Case 15746:  ControlIdToName = "Count Records"
        Case 15971:  ControlIdToName = "SharePoint Site Recycle Bin"
        Case 15973:  ControlIdToName = "Relink Lists"
        Case 16200:  ControlIdToName = "Document Management Server"
        Case 16201:  ControlIdToName = "Republish Database"
        Case 16206:  ControlIdToName = "Next Week"
        Case 16641:  ControlIdToName = "Expand All"
        Case 16642:  ControlIdToName = "Collapse All"
        Case 16667:  ControlIdToName = "Add Action"
        Case 16669:  ControlIdToName = "Add Copy of Macro"
        Case 16670:  ControlIdToName = "Add RunMacro"
        Case 16671:  ControlIdToName = "Expand Actions"
        Case 16672:  ControlIdToName = "Collapse Actions"
        Case 16673:  ControlIdToName = "Move Up"
        Case 16674:  ControlIdToName = "Move Down"
        Case 16676:  ControlIdToName = "Make Group Block"
        Case 16677:  ControlIdToName = "Make Submacro Block"
        Case 16678:  ControlIdToName = "Add Else If"
        Case 16679:  ControlIdToName = "Add Else"
        Case 16680:  ControlIdToName = "Make If Block"
        Case 16757:  ControlIdToName = "Insert Above"
        Case 16758:  ControlIdToName = "Insert Below"
        Case 16759:  ControlIdToName = "Insert Left"
        Case 16760:  ControlIdToName = "Insert Right"
        Case 16761:  ControlIdToName = "Merge"
        Case 16762:  ControlIdToName = "Split Vertically"
        Case 16763:  ControlIdToName = "Split Horizontally"
        Case 17136:  ControlIdToName = "Delete Row"
        Case 18515:  ControlIdToName = "Short Text"
        Case 18516:  ControlIdToName = "Long Text"
        Case 18517:  ControlIdToName = "Rich Text"
        Case 18518:  ControlIdToName = "Number"
        Case 18519:  ControlIdToName = "Date  Time"
        Case 18520:  ControlIdToName = "Currency"
        Case 18521:  ControlIdToName = "Yes/No"
        Case 18522:  ControlIdToName = "Hyperlink"
        Case 18523:  ControlIdToName = "Attachment"
        Case 18541:  ControlIdToName = "Lookup  Relationship"
        Case 18895:  ControlIdToName = "Modify Expression"
        Case 18904:  ControlIdToName = "Insert Navigation Button"
        Case 19954:  ControlIdToName = "Convert to Local Table"
        Case 22226:  ControlIdToName = "Build Hyperlink"
        Case 22284:  ControlIdToName = "Form Properties"
        Case 22285:  ControlIdToName = "Report Properties"
        Case 22328:  ControlIdToName = "Paste as Fields"
        Case 22329:  ControlIdToName = "Text"
        Case 22330:  ControlIdToName = "Number"
        Case 22331:  ControlIdToName = "Currency"
        Case 22332:  ControlIdToName = "Yes/No"
        Case 22333:  ControlIdToName = "Date/Time"
        Case 22366:  ControlIdToName = "Delete Column"
        Case 25274:  ControlIdToName = "Duplicate"
        Case 26593:  ControlIdToName = "From Dynamics 365 (online)"
        Case 27341:  ControlIdToName = "Large Number"
        Case 27429:  ControlIdToName = "From SQL Server"
        Case 27430:  ControlIdToName = "From Azure Database"
        Case 27434:  ControlIdToName = "From Amazon Redshift"
        Case 30002:  ControlIdToName = "File"
        Case 30003:  ControlIdToName = "Edit"
        Case 30004:  ControlIdToName = "View"
        Case 30005:  ControlIdToName = "Insert"
        Case 30006:  ControlIdToName = "Format"
        Case 30007:  ControlIdToName = "Tools"
        Case 30009:  ControlIdToName = "Window"
        Case 30010:  ControlIdToName = "Help"
        Case 30012:  ControlIdToName = "Run"
        Case 30014:  ControlIdToName = "Records"
        Case 30015:  ControlIdToName = "Security"
        Case 30016:  ControlIdToName = "Query"
        Case 30018:  ControlIdToName = "Relationships"
        Case 30019:  ControlIdToName = "Object"
        Case 30031:  ControlIdToName = "Filter"
        Case 30035:  ControlIdToName = "Align"
        Case 30038:  ControlIdToName = "Add-ins"
        Case 30039:  ControlIdToName = "Go To"
        Case 30041:  ControlIdToName = "SQL Specific"
        Case 30042:  ControlIdToName = "Size"
        Case 30043:  ControlIdToName = "Horizontal Spacing"
        Case 30044:  ControlIdToName = "Vertical Spacing"
        Case 30045:  ControlIdToName = "Toolbars"
        Case 30094:  ControlIdToName = "Hyperlink"
        Case 30095:  ControlIdToName = "Send To"
        Case 30101:  ControlIdToName = "Get External Data"
        Case 30102:  ControlIdToName = "Change To"
        Case 30103:  ControlIdToName = "Sort"
        Case 30104:  ControlIdToName = "Analyze"
        Case 30106:  ControlIdToName = "Office Links"
        Case 30107:  ControlIdToName = "Database Objects"
        Case 30108:  ControlIdToName = "Zoom"
        Case 30109:  ControlIdToName = "Pages"
        Case 30110:  ControlIdToName = "Arrange Icons"
        Case 30176:  ControlIdToName = "Favorites"
        Case 30179:  ControlIdToName = "Source Code Control"
        Case 30233:  ControlIdToName = "Database Utilities"
        Case 30234:  ControlIdToName = "Bookmarks"
        Case 30253:  ControlIdToName = "PivotTable"
        Case 30328:  ControlIdToName = "Go"
        Case 30379:  ControlIdToName = "Macro"
        Case 30398:  ControlIdToName = "Toggle"
        Case 30439:  ControlIdToName = "Table"
        Case 30440:  ControlIdToName = "Show Panes"
        Case 30453:  ControlIdToName = "Subdatasheet"
        Case 30454:  ControlIdToName = "Diagram"
        Case 30464:  ControlIdToName = "Convert Database"
        Case 30465:  ControlIdToName = "Background"
        Case 30468:  ControlIdToName = "Online Collaboration"
        Case 30469:  ControlIdToName = "PivotChart"
        Case 30476:  ControlIdToName = "Add to Group"
        Case 31182:  ControlIdToName = "Sample Databases"
        Case 31194:  ControlIdToName = "Show Only the Top"
        Case 31195:  ControlIdToName = "Show Only the Bottom"
        Case 31220:  ControlIdToName = "Subform"
        Case 31318:  ControlIdToName = "Import"
        Case 31381:  ControlIdToName = "Layout"
        Case 31382:  ControlIdToName = "Gridlines"
        Case 31398:  ControlIdToName = "Total"
        Case 31427:  ControlIdToName = "More options"
        Case 31458:  ControlIdToName = "Export"
        Case 31574:  ControlIdToName = "Position"
        Case 31582:  ControlIdToName = "All Dates In Period"
        Case 31584:  ControlIdToName = "Category"
        Case 31585:  ControlIdToName = "Sort By"
        Case 31586:  ControlIdToName = "View By"
        Case 31593:  ControlIdToName = "Add to group"
        Case 31626:  ControlIdToName = "Insert"
        Case 31627:  ControlIdToName = "Merge/Split"
        Case 31629:  ControlIdToName = "Calculated Field"
        Case 34008:  ControlIdToName = "Date  Time Extended"
        Case 34268:  ControlIdToName = "Dataverse"
        Case 34269:  ControlIdToName = "Dataverse"
        Case 34678:  ControlIdToName = "Legacy browser control"
        Case 34679:  ControlIdToName = "New browser control"
        Case 49304:  ControlIdToName = "175%"
        Case 49460:  ControlIdToName = "125%"
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsNonAddableControl
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Manual override that forces a built-in control to export as a custom
'           : replica regardless of the runtime addability probe. Empty by default:
'           : classification is handled comprehensively by IsBuiltInControlAddable. This
'           : is an escape hatch for the rare case where the probe reports a control as
'           : addable but it does not round-trip correctly. Add Ids here only when a
'           : real problem is observed; keep it minimal.
'---------------------------------------------------------------------------------------
'
Public Function IsNonAddableControl(lngId As Long) As Boolean

    Select Case lngId
        ' (Intentionally empty - see header. Add force-replica overrides here if needed.)
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsBuiltInControlAddable
' Author    : Adam Waller
' Date      : 6/25/2026
' Purpose   : Return true when a built-in control can be recreated with
'           : CommandBarControls.Add(Type, Id) on this Access version. Caches results
'           : per (Type, Id) for the Access session, so each distinct control is probed
'           : at most once. This is the comprehensive, self-maintaining classifier that
'           : replaces the hand-maintained whitelist/blacklist.
'---------------------------------------------------------------------------------------
'
Public Function IsBuiltInControlAddable(lngType As Long, lngId As Long) As Boolean

    Dim strKey As String

    If m_dAddableCache Is Nothing Then Set m_dAddableCache = New Dictionary
    strKey = lngType & "|" & lngId
    If Not m_dAddableCache.Exists(strKey) Then
        m_dAddableCache.Add strKey, ProbeBuiltInControlAddable(lngType, lngId)
    End If
    IsBuiltInControlAddable = m_dAddableCache(strKey)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ResetAddableCache
' Author    : Adam Waller
' Date      : 6/25/2026
' Purpose   : Clear the addability probe cache. Only needed by tests; addability is
'           : stable for the life of an Access session in normal use.
'---------------------------------------------------------------------------------------
'
Public Sub ResetAddableCache()
    Set m_dAddableCache = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ProbeBuiltInControlAddable
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Test whether Controls.Add(Type, Id) succeeds for one built-in control.
'           : Used by DumpNonAddableControls and the test suite. Creates and deletes a
'           : temporary scratch command bar.
'---------------------------------------------------------------------------------------
'
Public Function ProbeBuiltInControlAddable(lngType As Long, lngId As Long) As Boolean

    Const strTempBar As String = "__VCS_NonAddableTestBar__"

    Dim cb As CommandBar

    ProbeBuiltInControlAddable = TryAddBuiltInControl(lngType, lngId, strTempBar, cb)
    If Not cb Is Nothing Then
        cb.Delete
        CatchAny eelNoError, vbNullString
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : DumpNonAddableControls
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Scan Application.CommandBars for built-in control Ids, attempt
'           : Controls.Add(Type, Id) for each, and write a log of non-addable Ids with
'           : paste-ready Case lines for IsNonAddableControl. Never called during export.
'---------------------------------------------------------------------------------------
'
Public Sub DumpNonAddableControls()

    Dim cb As CommandBar
    Dim dBuiltIn As Dictionary
    Dim dResults As Dictionary
    Dim dSorted As Dictionary
    Dim strPath As String
    Dim strOutput As String
    Dim lngNonAddable As Long
    Dim lngNew As Long

    Debug.Print "Scanning built-in CommandBar controls for .Add(Type, Id) addability..."

    Set dBuiltIn = New Dictionary
    LogUnhandledErrors
    On Error Resume Next
    For Each cb In Application.CommandBars
        CollectBuiltInControlsById cb.Controls, dBuiltIn
    Next cb

    Set dResults = TestBuiltInControlAddability(dBuiltIn)
    Set dSorted = SortDictionaryByKeys(dResults)
    strOutput = BuildNonAddableDumpOutput(dSorted, lngNonAddable, lngNew)

    strPath = BuildPath2(Options.GetExportFolder, "logs\NonAddableControls_" & Format(Now, "yyyymmdd_hhnnss") & ".txt")
    VerifyPath strPath
    WriteFile strOutput, strPath

    Debug.Print "Tested " & dSorted.Count & " built-in control Id(s); " & lngNonAddable & " non-addable, " & lngNew & " new."
    Debug.Print "Wrote " & strPath

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CollectBuiltInControlsById
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Recursive helper for DumpNonAddableControls. De-duplicates built-in
'           : controls by Id; dictionary value is the control Type (for .Add testing).
'---------------------------------------------------------------------------------------
'
Private Sub CollectBuiltInControlsById(ctls As CommandBarControls, dEntries As Dictionary)

    Dim ctl As CommandBarControl
    Dim lngId As Long

    LogUnhandledErrors
    On Error Resume Next

    For Each ctl In ctls
        Err.Clear
        If ctl.BuiltIn Then
            lngId = ctl.Id
            If CatchAny(eelNoError, vbNullString) Then GoTo NextControl
            If Not dEntries.Exists(lngId) Then dEntries.Add lngId, ctl.Type
        End If
        If TypeOf ctl Is CommandBarPopup Then
            CollectBuiltInControlsById ctl.Controls, dEntries
        End If
NextControl:
    Next ctl

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestBuiltInControlAddability
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Attempt Controls.Add(Type, Id) for each collected built-in control on a
'           : temporary scratch command bar. Returns detail + paste-ready Case line per Id.
'---------------------------------------------------------------------------------------
'
Private Function TestBuiltInControlAddability(dBuiltIn As Dictionary) As Dictionary

    Const strTempBar As String = "__VCS_NonAddableTestBar__"

    Dim dResults As Dictionary
    Dim varId As Variant
    Dim lngId As Long
    Dim lngType As Long
    Dim blnAddable As Boolean
    Dim strName As String
    Dim strDetail As String
    Dim strCaseLine As String
    Dim cb As CommandBar

    Set dResults = New Dictionary
    LogUnhandledErrors
    On Error Resume Next

    ' Remove any leftover scratch bar from a prior interrupted run.
    Application.CommandBars(strTempBar).Delete
    CatchAny eelNoError, vbNullString

    For Each varId In dBuiltIn.Keys
        lngId = CLng(varId)
        lngType = CLng(dBuiltIn(lngId))
        blnAddable = TryAddBuiltInControl(lngType, lngId, strTempBar, cb)
        strName = ControlIdToName(lngId)
        If Len(strName) = 0 Then strName = "Id " & lngId
        strDetail = "Id=" & lngId & vbTab & "Type=" & lngType & vbTab & "Addable=" & blnAddable & vbTab & "Name=" & strName
        If blnAddable Then
            strCaseLine = "    ' Id " & lngId & " is addable"
        Else
            strCaseLine = "    Case " & lngId & ":  IsNonAddableControl = True    ' " & strName
        End If
        dResults.Add lngId, strDetail & vbCrLf & strCaseLine
    Next varId

    If Not cb Is Nothing Then cb.Delete
    CatchAny eelNoError, vbNullString

    Set TestBuiltInControlAddability = dResults

End Function


'---------------------------------------------------------------------------------------
' Procedure : TryAddBuiltInControl
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Create or reuse a temporary command bar and attempt to add one built-in
'           : control. Returns true when .Add succeeds.
'---------------------------------------------------------------------------------------
'
Private Function TryAddBuiltInControl(lngType As Long, lngId As Long, strTempBar As String, ByRef cb As CommandBar) As Boolean

    Dim ctl As CommandBarControl

    LogUnhandledErrors
    On Error Resume Next

    If cb Is Nothing Then
        Set cb = Application.CommandBars.Add(strTempBar, , False, True)
        If CatchAny(eelWarning, "Unable to create temporary command bar for addability test", ModuleName & ".TryAddBuiltInControl") Then Exit Function
    Else
        Do While cb.Controls.Count > 0
            cb.Controls(1).Delete
        Loop
    End If

    Err.Clear
    Set ctl = cb.Controls.Add(lngType, lngId)
    TryAddBuiltInControl = (Err.Number = 0) And Not ctl Is Nothing
    If Not CatchAny(eelNoError, vbNullString) Then Err.Clear

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildNonAddableDumpOutput
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Build log file contents: non-addable controls, new entries not yet in
'           : IsNonAddableControl, and the full sorted test results.
'---------------------------------------------------------------------------------------
'
Private Function BuildNonAddableDumpOutput(dEntries As Dictionary, ByRef lngNonAddable As Long, ByRef lngNew As Long) As String

    Dim varId As Variant
    Dim strNonAddable As String
    Dim strNew As String
    Dim strAll As String
    Dim strBlock As String
    Dim blnAddable As Boolean

    lngNonAddable = 0
    lngNew = 0
    strNonAddable = "' --- Non-addable controls (paste these Case lines into IsNonAddableControl) ---" & vbCrLf
    strNew = "' --- New non-addable not yet in IsNonAddableControl ---" & vbCrLf
    strAll = "' --- All built-in controls tested ---" & vbCrLf

    For Each varId In dEntries.Keys
        strBlock = dEntries(varId) & vbCrLf
        strAll = strAll & strBlock
        blnAddable = (InStr(1, strBlock, "Addable=True") > 0)
        If Not blnAddable Then
            lngNonAddable = lngNonAddable + 1
            strNonAddable = strNonAddable & strBlock
            If Not IsNonAddableControl(CLng(varId)) Then
                lngNew = lngNew + 1
                strNew = strNew & strBlock
            End If
        End If
    Next varId

    BuildNonAddableDumpOutput = strNonAddable & vbCrLf & strNew & vbCrLf & strAll

End Function


'---------------------------------------------------------------------------------------
' Procedure : DumpCommandBarControlNames
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Scan Application.CommandBars and list every control Id with its live
'           : Caption. Used only to author/grow ControlIdToName — never called during
'           : export. Writes de-duplicated, Id-sorted output to the Immediate window
'           : and a log file. Emits paste-ready Case lines for Ids not yet in the table.
'---------------------------------------------------------------------------------------
'
Public Sub DumpCommandBarControlNames()

    Dim cb As CommandBar
    Dim dEntries As Dictionary
    Dim dSorted As Dictionary
    Dim strPath As String
    Dim strOutput As String
    Dim lngMissing As Long
    Dim lngNoCaption As Long

    Debug.Print "Scanning Application.CommandBars for control Ids..."

    Set dEntries = New Dictionary
    LogUnhandledErrors
    On Error Resume Next
    For Each cb In Application.CommandBars
        CollectCommandBarControls cb.Controls, dEntries, cb.Name
    Next cb

    Set dSorted = SortDictionaryByKeys(dEntries)
    strOutput = BuildDumpOutput(dSorted, lngMissing, lngNoCaption)

    strPath = BuildPath2(Options.GetExportFolder, "logs\CommandBarControlNames_" & Format(Now, "yyyymmdd_hhnnss") & ".txt")
    VerifyPath strPath
    WriteFile strOutput, strPath

    Debug.Print "Scanned " & dSorted.Count & " control Id(s); " & _
        lngMissing & " missing, " & lngNoCaption & " with no usable caption."
    Debug.Print "Wrote " & strPath

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CollectCommandBarControls
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Recursive helper for DumpCommandBarControlNames. De-duplicates by Id and
'           : stores two lines per control (detail + paste-ready Case line).
'---------------------------------------------------------------------------------------
'
Private Sub CollectCommandBarControls(ctls As CommandBarControls, dEntries As Dictionary, strBarName As String)

    Dim ctl As CommandBarControl
    Dim strCaption As String
    Dim strRawCaption As String
    Dim strName As String
    Dim strDetail As String
    Dim strCaseLine As String
    Dim lngId As Long
    Dim lngType As Long

    LogUnhandledErrors
    On Error Resume Next

    For Each ctl In ctls
        Err.Clear
        lngId = ctl.Id
        lngType = ctl.Type
        strRawCaption = ctl.Caption
        If CatchAny(eelNoError, vbNullString) Then GoTo NextControl

        ' Key by the numeric Id so SortDictionaryByKeys orders the dump by Id value.
        If Not dEntries.Exists(lngId) Then
            If IsCaptionReadable(strRawCaption) Then
                strCaption = SanitizeCommandBarCaption(strRawCaption)
            Else
                strCaption = vbNullString
            End If
            strName = ControlIdToName(lngId)
            If Len(strName) = 0 And Len(strCaption) > 0 Then strName = strCaption
            strDetail = strBarName & vbTab & "Id=" & lngId & vbTab & "Type=" & lngType & vbTab & "Caption=" & strRawCaption
            If Len(strName) > 0 Then
                strCaseLine = "    Case " & lngId & ":  ControlIdToName = """ & Replace(strName, """", "'") & """    ' " & strRawCaption
            Else
                strCaseLine = "    ' Id " & lngId & " has no caption"
            End If
            dEntries.Add lngId, strDetail & vbCrLf & strCaseLine
        End If
        If TypeOf ctl Is CommandBarPopup Then
            CollectCommandBarControls ctl.Controls, dEntries, strBarName
        End If
NextControl:
    Next ctl

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildDumpOutput
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Build the log file contents in three sections:
'           :   1. Missing — readable caption but not in ControlIdToName (actionable;
'           :      paste the Case lines into ControlIdToName).
'           :   2. No Caption — empty or non-ASCII captions that cannot be named
'           :      (informational only, nothing to do).
'           :   3. All Controls — the full sorted list for reference.
'           : dEntries is expected to already be sorted by Id (SortDictionaryByKeys).
'---------------------------------------------------------------------------------------
'
Private Function BuildDumpOutput(dEntries As Dictionary, ByRef lngMissing As Long, ByRef lngNoCaption As Long) As String

    Dim varId As Variant
    Dim strMissing As String
    Dim strNoCaption As String
    Dim strAll As String
    Dim strBlock As String

    lngMissing = 0
    lngNoCaption = 0
    strMissing = "' --- Missing from ControlIdToName (paste these Case lines) ---" & vbCrLf
    strNoCaption = "' --- No usable caption (informational only) ---" & vbCrLf
    strAll = "' --- All controls by Id ---" & vbCrLf

    For Each varId In dEntries.Keys
        strBlock = dEntries(varId) & vbCrLf
        strAll = strAll & strBlock
        If Len(ControlIdToName(CLng(varId))) = 0 Then
            If InStr(1, strBlock, "    Case ") > 0 Then
                lngMissing = lngMissing + 1
                strMissing = strMissing & strBlock
            Else
                lngNoCaption = lngNoCaption + 1
                strNoCaption = strNoCaption & strBlock
            End If
        End If
    Next varId

    BuildDumpOutput = strMissing & vbCrLf & strNoCaption & vbCrLf & strAll

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsCaptionReadable
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Returns false unless every character in a live CommandBar caption is
'           : printable 7-bit ASCII. VBA source has limited Unicode support, and the
'           : table is only a human/agent breadcrumb, so captions carrying control
'           : characters or exotic glyphs (often uninitialized Access memory) are
'           : listed as comment-only entries rather than bogus Case lines.
'---------------------------------------------------------------------------------------
'
Private Function IsCaptionReadable(strCaption As String) As Boolean

    Dim lngPos As Long
    Dim lngCode As Long

    If Len(strCaption) = 0 Then Exit Function

    For lngPos = 1 To Len(strCaption)
        lngCode = AscW(Mid$(strCaption, lngPos, 1))
        ' Accept only printable 7-bit ASCII (space through tilde).
        If lngCode < 32 Or lngCode > 126 Then Exit Function
    Next lngPos

    IsCaptionReadable = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeCommandBarCaption
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Strip accelerator markers and trailing ellipses from a caption when
'           : generating paste-ready Case lines from a live scan.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeCommandBarCaption(strCaption As String) As String

    SanitizeCommandBarCaption = Trim$(MultiReplace(strCaption, _
        "&", vbNullString, _
        "...", vbNullString, _
        ":", vbNullString))

End Function
