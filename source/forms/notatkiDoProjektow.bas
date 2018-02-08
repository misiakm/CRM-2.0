Version =20
VersionRequired =20
PublishOption =1
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4793
    DatasheetFontHeight =11
    ItemSuffix =38
    Right =25575
    Bottom =12540
    DatasheetGridlinesColor =14806254
        0xe8ed93f43210e540
    End
    RecordSource ="SELECT notatkiDoProjektow.*, IIf(IsNull([dataDodania])=False,UCase([imie] & \" \""
        " & [nazwisko]) & \"   \" & [dataDodania],\"Nowa notatka:\") AS opis FROM uzytkow"
        "nicy RIGHT JOIN notatkiDoProjektow ON uzytkownicy.ID = notatkiDoProjektow.ktoDod"
        "al; "
    DatasheetFontName ="Calibri"
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
            BackStyle =0
            TextFontCharSet =238
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
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
            AddColon = NotDefault
            FELineBreak = NotDefault
            TextFontCharSet =238
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
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
            AddColon = NotDefault
            TextFontCharSet =238
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
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
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
            CanGrow = NotDefault
            Height =0
            BackColor =15849926
            Name ="NagłówekFormularza"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
            CanGrow = NotDefault
            Height =930
            Name ="Szczegóły"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =30
                    Width =4755
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    BorderColor =10921638
                    Name ="opis"
                    ControlSource ="opis"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =30
                    LayoutCachedWidth =4785
                    LayoutCachedHeight =255
                    LayoutGroup =1
                    ForeTint =100.0
                    GroupTable =1
                End
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =30
                    Top =255
                    Width =4755
                    Height =675
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="notatka"
                    ControlSource ="notatka"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"Nowa notatka\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =30
                    LayoutCachedTop =255
                    LayoutCachedWidth =4785
                    LayoutCachedHeight =930
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="ktoDodal"
                    ControlSource ="ktoDodal"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [uzytkownicy].[ID], [uzytkownicy].[imie], [uzytkownicy].[nazwisko] FROM u"
                        "zytkownicy; "
                    ColumnWidths ="0;1441;1441"
                    DefaultValue ="getUser()"
                    BaseInfo ="\"SELECT [uzytkownicy].[ID], [uzytkownicy].[imie], [uzytkownicy].[nazwisko] FROM"
                        " uzytkownicy; \";\"uzytkownicy\";\"\";\"ID\";\"imie\";\"PrimaryKey\""
                    GridlineColor =10921638

                    LayoutCachedWidth =1701
                    LayoutCachedHeight =315
                End
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="projekt"
                    ControlSource ="projekt"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [projekty].[ID], [projekty].[nazwaProjektu] FROM projekty; "
                    ColumnWidths ="0;1441"
                    BaseInfo ="\"SELECT [projekty].[ID], [projekty].[nazwaProjektu] FROM projekty; \";\"projekt"
                        "y\";\"\";\"ID\";\"nazwaProjektu\";\"PrimaryKey\""
                    GridlineColor =10921638

                    LayoutCachedWidth =1701
                    LayoutCachedHeight =315
                End
            End
        End
            Height =0
            Name ="StopkaFormularza"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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

Private Sub notatka_AfterUpdate()
   Me.Requery
End Sub
