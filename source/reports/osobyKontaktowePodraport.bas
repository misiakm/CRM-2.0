Version =20
VersionRequired =20
PublishOption =1
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =2895
    DatasheetFontHeight =11
    ItemSuffix =68
    Right =20550
    Bottom =10125
    DatasheetGridlinesColor =14806254
        0x762013700f10e540
    End
    RecordSource ="SELECT * FROM (Select p.ID As IDProjektu, nazwaProjektu, imie & \" \" & nazwisko"
        " As osobaKontaktowa, true As czyDecyzyjna, ok.*  from projekty as p  left join o"
        "sobyKontaktowe As ok on ok.ID = p.osobaDecyzyjna    Union All    Select p.ID As "
        "IDProjektu,nazwaProjektu, imie & \" \" & nazwisko As osobaKontaktowa, false As c"
        "zyDecyzyjna, ok.*  FROM osobyKontaktowe AS ok RIGHT JOIN (projekty AS p LEFT JOI"
        "N osobyKontaktoweDoProjektow AS okdp ON p.ID = okdp.projekt) ON ok.ID = okdp.oso"
        "baKontaktowa)  AS t WHERE (((t.osobaKontaktowa)<>\" \")); "
    DatasheetFontName ="Calibri"
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            BorderLineStyle =0
            Width =1701
            BorderThemeColorIndex =0
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
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
            GroupHeader = NotDefault
            ControlSource ="IDProjektu"
        End
            ControlSource ="czyDecyzyjna"
        End
            Height =0
            Name ="SekcjaNagłówkaStrony"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =450
            BackColor =4144959
            Name ="NagłówekGrupy0"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Width =2841
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =4144959
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="nazwaProjektu"
                    ControlSource ="nazwaProjektu"
                    GridlineColor =10921638

                    LayoutCachedWidth =2841
                    LayoutCachedHeight =405
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
            KeepTogether = NotDefault
            CanShrink = NotDefault
            Height =2055
            Name ="Szczegóły"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Width =2895
                    Height =390
                    FontSize =13
                    BackColor =16119285
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tekst47"
                    ControlSource ="osobaKontaktowa"
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000fffa9300000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0063007a00790044006500630079007a0079006a006e0061005d003d005400 ,
                        0x72007500650000000000
                    End
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    GridlineWidthTop =0

                    LayoutCachedWidth =2895
                    LayoutCachedHeight =390
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                        0x01000100000001000000000000000100000000000000fffa9300130000005b00 ,
                        0x63007a00790044006500630079007a0079006a006e0061005d003d0054007200 ,
                        0x75006500000000000000000000000000000000000000000000
                    End
                    GroupTable =2
                End
                    CanShrink = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Top =705
                    Width =2895
                    Height =315
                    TabIndex =1
                    BackColor =16119285
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="mail"
                    ControlSource ="mail"
                    OnClick ="[Event Procedure]"
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000fffa9300000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0063007a00790044006500630079007a0079006a006e0061005d003d005400 ,
                        0x72007500650000000000
                    End
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    GridlineWidthTop =0

                    LayoutCachedTop =705
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =1020
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                        0x01000100000001000000000000000100000000000000fffa9300130000005b00 ,
                        0x63007a00790044006500630079007a0079006a006e0061005d003d0054007200 ,
                        0x75006500000000000000000000000000000000000000000000
                    End
                    GroupTable =2
                End
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Top =390
                    Width =2895
                    Height =315
                    TabIndex =2
                    BackColor =16119285
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tekst49"
                    ControlSource ="opis"
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000fffa9300000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0063007a00790044006500630079007a0079006a006e0061005d003d005400 ,
                        0x72007500650000000000
                    End
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    GridlineWidthTop =0

                    LayoutCachedTop =390
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =705
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                        0x01000100000001000000000000000100000000000000fffa9300130000005b00 ,
                        0x63007a00790044006500630079007a0079006a006e0061005d003d0054007200 ,
                        0x75006500000000000000000000000000000000000000000000
                    End
                    GroupTable =2
                End
                    CanShrink = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Top =1020
                    Width =2895
                    Height =315
                    TabIndex =3
                    BackColor =16119285
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="mail2"
                    ControlSource ="mail2"
                    OnClick ="[Event Procedure]"
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000fffa9300000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0063007a00790044006500630079007a0079006a006e0061005d003d005400 ,
                        0x72007500650000000000
                    End
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    GridlineWidthTop =0

                    LayoutCachedTop =1020
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                        0x01000100000001000000000000000100000000000000fffa9300130000005b00 ,
                        0x63007a00790044006500630079007a0079006a006e0061005d003d0054007200 ,
                        0x75006500000000000000000000000000000000000000000000
                    End
                    GroupTable =2
                End
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Top =1335
                    Width =2895
                    Height =315
                    TabIndex =4
                    BackColor =16119285
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tekst51"
                    ControlSource ="telefon"
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000fffa9300000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0063007a00790044006500630079007a0079006a006e0061005d003d005400 ,
                        0x72007500650000000000
                    End
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    GridlineWidthTop =0

                    LayoutCachedTop =1335
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =1650
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                        0x01000100000001000000000000000100000000000000fffa9300130000005b00 ,
                        0x63007a00790044006500630079007a0079006a006e0061005d003d0054007200 ,
                        0x75006500000000000000000000000000000000000000000000
                    End
                    GroupTable =2
                End
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Top =1650
                    Width =2895
                    Height =315
                    TabIndex =5
                    BackColor =16119285
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tekst52"
                    ControlSource ="telefon2"
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000fffa9300000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0063007a00790044006500630079007a0079006a006e0061005d003d005400 ,
                        0x72007500650000000000
                    End
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    GridlineWidthTop =0

                    LayoutCachedTop =1650
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =1965
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                        0x01000100000001000000000000000100000000000000fffa9300130000005b00 ,
                        0x63007a00790044006500630079007a0079006a006e0061005d003d0054007200 ,
                        0x75006500000000000000000000000000000000000000000000
                    End
                    GroupTable =2
                End
            End
        End
            Height =0
            Name ="SekcjaStopkiStrony"
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


Private Sub mail_Click()
   Call konstruujMaila(Me.mail)
End Sub

Private Sub mail2_Click()
   Call konstruujMaila(Me.mail2)
End Sub

Sub konstruujMaila(adres As String)
   Dim s As String
   s = "mailto:" & adres
   Application.FollowHyperlink s
End Sub
