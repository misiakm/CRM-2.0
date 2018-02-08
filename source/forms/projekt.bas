Version =20
VersionRequired =20
PublishOption =1
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =18086
    DatasheetFontHeight =11
    ItemSuffix =179
    Left =-15
    Top =660
    Right =18780
    Bottom =9225
    DatasheetGridlinesColor =14806254
    Filter ="Idprojektu = 5"
        0x08e507123410e540
    End
    RecordSource ="SELECT projekty.*, Organizacje.*, [uzytkownicy].[imie] & \" \" & [uzytkownicy].["
        "nazwisko] AS wlasciciel, [osobyKontaktowe].[imie] & \" \" & [osobyKontaktowe].[n"
        "azwisko] AS osDecyzyjna, osobyKontaktowe.*, kategorieOrganizacji.kategoria, typy"
        "Organizacij.typ, projekty.ID AS IDProjektu FROM uzytkownicy RIGHT JOIN (typyOrga"
        "nizacij RIGHT JOIN (osobyKontaktowe RIGHT JOIN (kategorieOrganizacji RIGHT JOIN "
        "(Organizacje INNER JOIN projekty ON Organizacje.ID = projekty.organizacja) ON ka"
        "tegorieOrganizacji.ID = Organizacje.kategoriaFirmy) ON osobyKontaktowe.ID = proj"
        "ekty.osobaDecyzyjna) ON typyOrganizacij.ID = Organizacje.typFirmy) ON uzytkownic"
        "y.ID = projekty.ktoDodal; "
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
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            TextFontCharSet =238
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
            Height =1814
            Name ="NagłówekFormularza"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2655
                    Top =120
                    Width =7431
                    Height =510
                    ColumnOrder =0
                    FontSize =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="nazwaProjektu"
                    ControlSource ="nazwaProjektu"
                    GridlineColor =10921638

                    LayoutCachedLeft =2655
                    LayoutCachedTop =120
                    LayoutCachedWidth =10086
                    LayoutCachedHeight =630
                End
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2659
                    Top =751
                    Width =7416
                    Height =405
                    ColumnOrder =1
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="wartosc"
                    ControlSource ="wartosc"
                    Format ="#,##0.00\" zł\";-#,##0.00\" zł\""
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =2659
                    LayoutCachedTop =751
                    LayoutCachedWidth =10075
                    LayoutCachedHeight =1156
                    CurrencySymbol ="zł"
                End
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =2659
                    Top =1261
                    Width =3576
                    Height =315
                    ColumnOrder =2
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="etapWLejku"
                    ControlSource ="etapWLejku"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [lejki].[ID], [lejki].[nazwa] FROM lejki; "
                    ColumnWidths ="0;1440"
                    BaseInfo ="\"SELECT [lejki].[ID], [lejki].[nazwa] FROM lejki; \";\"lejki\";\"\";\"ID\";\"na"
                        "zwa\";\"PrimaryKey\""
                    GridlineColor =10921638

                    LayoutCachedLeft =2659
                    LayoutCachedTop =1261
                    LayoutCachedWidth =6235
                    LayoutCachedHeight =1576
                End
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =2190
                    Height =465
                    TabIndex =1
                    ForeColor =16777215
                    Name ="btnSukces"
                    Caption ="Sukces"
                    OnClick ="[Event Procedure]"
                    GroupTable =3
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =585
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =5026082
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =5167783
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =5026082
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =5026082
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =3
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                    OverlapFlags =85
                    Left =120
                    Top =645
                    Width =2190
                    Height =465
                    TabIndex =2
                    ForeColor =16777215
                    Name ="btnPorazka"
                    Caption ="Porażka"
                    OnClick ="[Event Procedure]"
                    GroupTable =3
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =645
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =1110
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =1643706
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =2366701
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =1643706
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =1643706
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =3
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
            CanGrow = NotDefault
            Height =7381
            BackColor =16514043
            Name ="Szczegóły"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =3750
                    Top =60
                    Width =5160
                    Height =7321
                    BorderColor =16382457
                    Name ="Pole164"
                    GridlineColor =10921638
                    LayoutCachedLeft =3750
                    LayoutCachedTop =60
                    LayoutCachedWidth =8910
                    LayoutCachedHeight =7381
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =113
                    Top =4081
                    Width =3450
                    Height =3286
                    BorderColor =16382457
                    Name ="Pole160"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedTop =4081
                    LayoutCachedWidth =3563
                    LayoutCachedHeight =7367
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =120
                    Top =60
                    Width =3450
                    Height =3916
                    BorderColor =16382457
                    Name ="Pole59"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =3570
                    LayoutCachedHeight =3976
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                    OverlapFlags =215
                    Left =233
                    Top =180
                    Width =2145
                    Height =390
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =4144959
                    Name ="Etykieta57"
                    Caption ="Osoba decyzyjna"
                    GridlineColor =10921638
                    LayoutCachedLeft =233
                    LayoutCachedTop =180
                    LayoutCachedWidth =2378
                    LayoutCachedHeight =570
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =285
                    Top =1755
                    Width =3120
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="mail"
                    ControlSource ="mail"
                    GroupTable =4
                    GridlineColor =10921638

                    LayoutCachedLeft =285
                    LayoutCachedTop =1755
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =2070
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =4
                End
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =285
                    Top =2430
                    Width =3120
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="mail2"
                    ControlSource ="mail2"
                    GroupTable =4
                    GridlineColor =10921638

                    LayoutCachedLeft =285
                    LayoutCachedTop =2430
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =2745
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =4
                End
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =285
                    Top =3105
                    Width =1500
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="telefon"
                    ControlSource ="telefon"
                    GroupTable =4
                    GridlineColor =10921638

                    LayoutCachedLeft =285
                    LayoutCachedTop =3105
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =3420
                    RowStart =7
                    RowEnd =7
                    LayoutGroup =2
                    GroupTable =4
                End
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1845
                    Top =3105
                    Width =1560
                    Height =315
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="telefon2"
                    ControlSource ="telefon2"
                    GroupTable =4
                    GridlineColor =10921638

                    LayoutCachedLeft =1845
                    LayoutCachedTop =3105
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =3420
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =4
                End
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =285
                    Top =1035
                    Width =1500
                    Height =390
                    FontSize =13
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="imie"
                    ControlSource ="imie"
                    GroupTable =4
                    GridlineColor =10921638

                    LayoutCachedLeft =285
                    LayoutCachedTop =1035
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =1425
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =4
                End
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1845
                    Top =1035
                    Width =1560
                    Height =390
                    FontSize =13
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="nazwisko"
                    ControlSource ="nazwisko"
                    GroupTable =4
                    GridlineColor =10921638

                    LayoutCachedLeft =1845
                    LayoutCachedTop =1035
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =1425
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =4
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =285
                    Top =750
                    Width =1500
                    Height =225
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta83"
                    Caption ="IMIĘ"
                    GroupTable =4
                    GridlineColor =10921638
                    LayoutCachedLeft =285
                    LayoutCachedTop =750
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =975
                    LayoutGroup =2
                    ForeTint =100.0
                    GroupTable =4
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =1845
                    Top =750
                    Width =1560
                    Height =225
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta88"
                    Caption ="NAZWISKO"
                    GroupTable =4
                    GridlineColor =10921638
                    LayoutCachedLeft =1845
                    LayoutCachedTop =750
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =975
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeTint =100.0
                    GroupTable =4
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =285
                    Top =1485
                    Width =1500
                    Height =210
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta90"
                    Caption ="E-MAIL"
                    GroupTable =4
                    GridlineColor =10921638
                    LayoutCachedLeft =285
                    LayoutCachedTop =1485
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =1695
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    ForeTint =100.0
                    GroupTable =4
                End
                    Left =1845
                    Top =1485
                    Width =1560
                    Height =210
                    Name ="PustaKomórka93"
                    GroupTable =4
                    GridlineColor =10921638
                    LayoutCachedLeft =1845
                    LayoutCachedTop =1485
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =1695
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =4
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =285
                    Top =2130
                    Width =1500
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta94"
                    Caption ="E-MAIL 2"
                    GroupTable =4
                    GridlineColor =10921638
                    LayoutCachedLeft =285
                    LayoutCachedTop =2130
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =2370
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    ForeTint =100.0
                    GroupTable =4
                End
                    Left =1845
                    Top =2130
                    Width =1560
                    Name ="PustaKomórka98"
                    GroupTable =4
                    GridlineColor =10921638
                    LayoutCachedLeft =1845
                    LayoutCachedTop =2130
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =2370
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =4
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =285
                    Top =2805
                    Width =1500
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta99"
                    Caption ="TELEFON"
                    GroupTable =4
                    GridlineColor =10921638
                    LayoutCachedLeft =285
                    LayoutCachedTop =2805
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =3045
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =2
                    ForeTint =100.0
                    GroupTable =4
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =1845
                    Top =2805
                    Width =1560
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta104"
                    Caption ="TELEFON 2"
                    GroupTable =4
                    GridlineColor =10921638
                    LayoutCachedLeft =1845
                    LayoutCachedTop =2805
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =3045
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeTint =100.0
                    GroupTable =4
                End
                    OverlapFlags =215
                    Left =2998
                    Top =113
                    Width =520
                    Height =451
                    ForeColor =4210752
                    Name ="btnEdytujOsobeDecyzyjna"
                    GridlineColor =10921638
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4af9b17d4a78b17d4a18 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a81b17d4affb17d4af3 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a27b17d4af6b17d4a03 ,
                        0xb17d4ab7b17d4a6c000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0cb17d4ab7 ,
                        0xb17d4affb17d4affb17d4a7b0000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a87 ,
                        0xb17d4affb17d4affb17d4affb17d4a8700000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb17d4a8db17d4affb17d4affb17d4affb17d4a93000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a90b17d4affb17d4affb17d4aabb17d4a12b17d4a0300000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a96b17d4aabb17d4a15b17d4acfb17d4aa500000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a12b17d4acfb17d4affb17d4af000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a03b17d4aa2b17d4afcb17d4a2a00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =2998
                    LayoutCachedTop =113
                    LayoutCachedWidth =3518
                    LayoutCachedHeight =564
                    Gradient =0
                    BackColor =14136213
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =-2
                    HoverThemeColorIndex =-1
                    PressedColor =-2
                    PressedThemeColorIndex =-1
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                    OverlapFlags =215
                    Left =227
                    Top =4195
                    Width =2145
                    Height =390
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =4144959
                    Name ="Etykieta110"
                    Caption ="Organizacja"
                    GridlineColor =10921638
                    LayoutCachedLeft =227
                    LayoutCachedTop =4195
                    LayoutCachedWidth =2372
                    LayoutCachedHeight =4585
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                    OverlapFlags =215
                    Left =285
                    Top =3540
                    Width =3120
                    Height =315
                    TabIndex =9
                    ForeColor =16777215
                    Name ="Polecenie111"
                    Caption ="Pokaż osoby kontaktowe"
                    GroupTable =4
                    TopPadding =90
                    BottomPadding =105
                    GridlineColor =10921638

                    LayoutCachedLeft =285
                    LayoutCachedTop =3540
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =3855
                    RowStart =8
                    RowEnd =8
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =5026082
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =5167783
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =5026082
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =5026082
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =4
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =6
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =6
                    Overlaps =1
                End
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =300
                    Top =5055
                    Width =3120
                    Height =375
                    FontSize =13
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="nazwaFirmy"
                    ControlSource ="nazwaFirmy"
                    GroupTable =14
                    LeftPadding =90
                    RightPadding =105
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =5055
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =5430
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    GroupTable =14
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =300
                    Top =4770
                    Width =3120
                    Height =225
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta117"
                    Caption ="NAZWA"
                    GroupTable =14
                    LeftPadding =90
                    RightPadding =105
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =4770
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =4995
                    LayoutGroup =3
                    ForeTint =100.0
                    GroupTable =14
                End
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =300
                    Top =5790
                    Width =3120
                    Height =330
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="kategoria"
                    ControlSource ="kategoria"
                    GroupTable =14
                    LeftPadding =90
                    RightPadding =105
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =5790
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =6120
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =3
                    GroupTable =14
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =300
                    Top =5490
                    Width =3120
                    Height =240
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta119"
                    Caption ="KATEGORIA"
                    GroupTable =14
                    LeftPadding =90
                    RightPadding =105
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =5490
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =5730
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    ForeTint =100.0
                    GroupTable =14
                End
                    OverlapFlags =215
                    TextAlign =1
                    Left =300
                    Top =6180
                    Width =3120
                    Height =225
                    FontSize =9
                    BorderColor =8355711
                    Name ="Etykieta131"
                    Caption ="TYP FIRMY"
                    GroupTable =14
                    LeftPadding =90
                    RightPadding =105
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =6180
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =6405
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =3
                    ForeTint =100.0
                    GroupTable =14
                End
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =300
                    Top =6465
                    Width =3120
                    Height =330
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="typ"
                    ControlSource ="typ"
                    GroupTable =14
                    LeftPadding =90
                    RightPadding =105
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =6465
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =6795
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =3
                    GroupTable =14
                End
                    OverlapFlags =215
                    Left =300
                    Top =6915
                    Width =3120
                    Height =315
                    TabIndex =13
                    ForeColor =16777215
                    Name ="Polecenie149"
                    Caption ="Pokaż więcej szczegółów"
                    GroupTable =14
                    LeftPadding =90
                    TopPadding =90
                    RightPadding =105
                    BottomPadding =105
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =6915
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =7230
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =3
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =5026082
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =5167783
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =5026082
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =5026082
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =14
                    WebImagePaddingLeft =6
                    WebImagePaddingTop =6
                    WebImagePaddingRight =6
                    WebImagePaddingBottom =6
                    Overlaps =1
                End
                    OverlapFlags =215
                    Left =2891
                    Top =4195
                    Width =520
                    Height =451
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Polecenie161"
                    GridlineColor =10921638
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4af9b17d4a78b17d4a18 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a81b17d4affb17d4af3 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a27b17d4af6b17d4a03 ,
                        0xb17d4ab7b17d4a6c000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0cb17d4ab7 ,
                        0xb17d4affb17d4affb17d4a7b0000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a87 ,
                        0xb17d4affb17d4affb17d4affb17d4a8700000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb17d4a8db17d4affb17d4affb17d4affb17d4a93000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a90b17d4affb17d4affb17d4aabb17d4a12b17d4a0300000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a96b17d4aabb17d4a15b17d4acfb17d4aa500000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a12b17d4acfb17d4affb17d4af000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a03b17d4aa2b17d4afcb17d4a2a00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =2891
                    LayoutCachedTop =4195
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =4646
                    Gradient =0
                    BackColor =14136213
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =-2
                    HoverThemeColorIndex =-1
                    PressedColor =-2
                    PressedThemeColorIndex =-1
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =3758
                    Top =632
                    Width =5130
                    Height =6630
                    TabIndex =2
                    BorderColor =10921638
                    Name ="podformularz"
                    SourceObject ="Form.notatkiDoProjektow"
                    LinkChildFields ="projekt"
                    LinkMasterFields ="IDProjektu"
                    GridlineColor =10921638

                    LayoutCachedLeft =3758
                    LayoutCachedTop =632
                    LayoutCachedWidth =8888
                    LayoutCachedHeight =7262
                End
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =3815
                    Top =170
                    Width =2550
                    Height =390
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =4144959
                    Name ="lbNotatki"
                    Caption ="Notatki"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =3815
                    LayoutCachedTop =170
                    LayoutCachedWidth =6365
                    LayoutCachedHeight =560
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                    OverlapFlags =215
                    TextAlign =2
                    Left =6424
                    Top =170
                    Width =2370
                    Height =390
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =4144959
                    Name ="lbZadania"
                    Caption ="Zadania"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =6424
                    LayoutCachedTop =170
                    LayoutCachedWidth =8794
                    LayoutCachedHeight =560
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =9079
                    Top =682
                    Width =8430
                    Height =6480
                    TabIndex =14
                    BorderColor =10921638
                    Name ="podformularzLista"
                    SourceObject ="Form.osobyKontaktowe"
                    GridlineColor =10921638

                    LayoutCachedLeft =9079
                    LayoutCachedTop =682
                    LayoutCachedWidth =17509
                    LayoutCachedHeight =7162
                End
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =8955
                    Top =60
                    Width =8625
                    Height =7201
                    BorderColor =16382457
                    Name ="Pole177"
                    GridlineColor =10921638
                    LayoutCachedLeft =8955
                    LayoutCachedTop =60
                    LayoutCachedWidth =17580
                    LayoutCachedHeight =7261
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                    OverlapFlags =215
                    Left =9075
                    Top =173
                    Width =2145
                    Height =390
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =4144959
                    Name ="Etykieta178"
                    Caption ="Osoby kontaktowe"
                    GridlineColor =10921638
                    LayoutCachedLeft =9075
                    LayoutCachedTop =173
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =563
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
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

Private Sub btnPorazka_Click()
   Call zmienStatus(2)
End Sub

Private Sub btnSukces_Click()
   Call zmienStatus(3)
End Sub

Private Sub lbNotatki_Click()
   Call zmienPodformularz("notatki")
End Sub

Private Sub lbZadania_Click()
   Call zmienPodformularz("zadania")
End Sub

Sub zmienPodformularz(zadaniaNotatki As String)
   Me.lbZadania.FontUnderline = zadaniaNotatki = "zadania"
   Me.lbNotatki.FontUnderline = zadaniaNotatki = "notatki"
   Me.podformularz.SourceObject = IIf(zadaniaNotatki = "zadania", "zadaniaKrotkie", "notatkiDoProjektow")
End Sub

Sub zmienStatus(status As Integer)
   Dim sql As String
   sql = "Update projekty set status = " & status & " where ID = " & Me.IDProjektu
   Call runSQL(sql)
   Call Form_panel.btnSzanseSprzedazy_Click
End Sub
