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
    Width =18885
    DatasheetFontHeight =11
    ItemSuffix =70
    Right =20550
    Bottom =12285
    DatasheetGridlinesColor =14806254
    RecordSource ="TRANSFORM First(\"X\") AS Wyr1 SELECT projekty.nazwaProjektu, projekty.etapWLejk"
        "u FROM lejki LEFT JOIN projekty ON lejki.ID = projekty.etapWLejku GROUP BY proje"
        "kty.nazwaProjektu, projekty.etapWLejku ORDER BY projekty.nazwaProjektu, lejki.ID"
        " PIVOT lejki.ID; "
    DatasheetFontName ="Calibri"
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
            Height =0
            BackColor =15849926
            Name ="NagłówekFormularza"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
            Height =396
            Name ="Szczegóły"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =2970
                    Height =315
                    FontWeight =700
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="nazwaProjektu"
                    ControlSource ="nazwaProjektu"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =345
                    LayoutGroup =1
                    GroupTable =1
                End
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3030
                    Top =30
                    Width =2268
                    Height =315
                    FontSize =10
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="lejek1"
                        0x01000000a0000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000d1eaf0000100000000000000100000001f0000000100000000000000 ,
                        0xa4d5e20000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003d00310000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003e00310000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3030
                    LayoutCachedTop =30
                    LayoutCachedWidth =5298
                    LayoutCachedHeight =345
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                        0x01000200000001000000000000000100000000000000d1eaf0000e0000005b00 ,
                        0x650074006100700057004c0065006a006b0075005d003d003100000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000a4 ,
                        0xd5e2000e0000005b00650074006100700057004c0065006a006b0075005d003e ,
                        0x003100000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5295
                    Top =30
                    Width =2265
                    Height =315
                    FontSize =10
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="lejek2"
                        0x01000000a0000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000d1eaf0000100000000000000100000001f0000000100000000000000 ,
                        0xa4d5e20000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003d00320000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003e00320000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5295
                    LayoutCachedTop =30
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =345
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                        0x01000200000001000000000000000100000000000000d1eaf0000e0000005b00 ,
                        0x650074006100700057004c0065006a006b0075005d003d003200000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000a4 ,
                        0xd5e2000e0000005b00650074006100700057004c0065006a006b0075005d003e ,
                        0x003200000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7560
                    Top =30
                    Width =2265
                    Height =315
                    FontSize =10
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="lejek3"
                        0x01000000a0000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000d1eaf0000100000000000000100000001f0000000100000000000000 ,
                        0xa4d5e20000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003d00320000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003e00320000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7560
                    LayoutCachedTop =30
                    LayoutCachedWidth =9825
                    LayoutCachedHeight =345
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                        0x01000200000001000000000000000100000000000000d1eaf0000e0000005b00 ,
                        0x650074006100700057004c0065006a006b0075005d003d003200000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000a4 ,
                        0xd5e2000e0000005b00650074006100700057004c0065006a006b0075005d003e ,
                        0x003200000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9825
                    Top =30
                    Width =2265
                    Height =315
                    FontSize =10
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="lejek4"
                        0x01000000a0000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000d1eaf0000100000000000000100000001f0000000100000000000000 ,
                        0xa4d5e20000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003d00320000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003e00320000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =9825
                    LayoutCachedTop =30
                    LayoutCachedWidth =12090
                    LayoutCachedHeight =345
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                        0x01000200000001000000000000000100000000000000d1eaf0000e0000005b00 ,
                        0x650074006100700057004c0065006a006b0075005d003d003200000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000a4 ,
                        0xd5e2000e0000005b00650074006100700057004c0065006a006b0075005d003e ,
                        0x003200000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12090
                    Top =30
                    Width =2265
                    Height =315
                    FontSize =10
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="lejek5"
                        0x01000000a0000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000d1eaf0000100000000000000100000001f0000000100000000000000 ,
                        0xa4d5e20000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003d00320000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003e00320000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =12090
                    LayoutCachedTop =30
                    LayoutCachedWidth =14355
                    LayoutCachedHeight =345
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                        0x01000200000001000000000000000100000000000000d1eaf0000e0000005b00 ,
                        0x650074006100700057004c0065006a006b0075005d003d003200000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000a4 ,
                        0xd5e2000e0000005b00650074006100700057004c0065006a006b0075005d003e ,
                        0x003200000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14355
                    Top =30
                    Width =2265
                    Height =315
                    FontSize =10
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="lejek6"
                        0x01000000a0000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000d1eaf0000100000000000000100000001f0000000100000000000000 ,
                        0xa4d5e20000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003d00320000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003e00320000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =14355
                    LayoutCachedTop =30
                    LayoutCachedWidth =16620
                    LayoutCachedHeight =345
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                        0x01000200000001000000000000000100000000000000d1eaf0000e0000005b00 ,
                        0x650074006100700057004c0065006a006b0075005d003d003200000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000a4 ,
                        0xd5e2000e0000005b00650074006100700057004c0065006a006b0075005d003e ,
                        0x003200000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16620
                    Top =30
                    Width =2265
                    Height =315
                    FontSize =10
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="lejek7"
                        0x01000000a0000000020000000100000000000000000000000f00000001000000 ,
                        0x00000000d1eaf0000100000000000000100000001f0000000100000000000000 ,
                        0xa4d5e20000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003d00320000000000 ,
                        0x5b00650074006100700057004c0065006a006b0075005d003e00320000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =16620
                    LayoutCachedTop =30
                    LayoutCachedWidth =18885
                    LayoutCachedHeight =345
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                        0x01000200000001000000000000000100000000000000d1eaf0000e0000005b00 ,
                        0x650074006100700057004c0065006a006b0075005d003d003200000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000a4 ,
                        0xd5e2000e0000005b00650074006100700057004c0065006a006b0075005d003e ,
                        0x003200000000000000000000000000000000000000000000
                    End
                    GroupTable =1
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

Private Sub Form_Load()
   Call zmienPrzyciskiLejki
End Sub

Sub zmienPrzyciskiLejki()
   Dim rec As Recordset, i As Integer, nazwa As String
   Set rec = CurrentDb.OpenRecordset("lejki")
   Do Until rec.EOF
      i = i + 1
      nazwa = IIf(Len(rec!nazwa) > 25, Left(rec!nazwa, 22) & "...", rec!nazwa)
      Me.Controls("lejek" & i) = nazwa
      Me.Controls("lejek" & i).Visible = True
      Call ustawFormatowanieWarunkowe(i)
      rec.MoveNext
   Loop
End Sub

Sub ustawFormatowanieWarunkowe(i As Integer)
   Dim tb As TextBox
   Set tb = Me.Controls("lejek" & i)
   tb.FormatConditions.Delete
   tb.FormatConditions.Add acExpression, , "[etapWLejku]>=" & i
   tb.FormatConditions(0).BackColor = "5026082"
   tb.FormatConditions.Add acExpression, , "[etapWLejku]<" & i
   tb.FormatConditions(1).BackColor = "1643706"
   tb.FormatConditions(1).ForeColor = 16777215
End Sub
