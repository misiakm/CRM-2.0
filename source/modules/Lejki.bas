Option Compare Database
Option Explicit

Sub zbierzLejki(f As Form, ByRef k As Collection)
   Dim c As Control, p As lejkiClass
   For Each c In f.Controls
      Set p = New lejkiClass
      If p.czyJestLejkiem(c) Then
         k.Add p
      End If
   Next c
End Sub

Sub zmienPrzyciskiLejki(f As Form)
   Dim rec As Recordset, i As Integer, nazwa As String
   Set rec = CurrentDb.OpenRecordset("lejki")
   Do Until rec.EOF
      i = i + 1
      nazwa = IIf(Len(rec!nazwa) > 25, Left(rec!nazwa, 22) & "...", rec!nazwa)
      f.Controls("lejek" & i) = nazwa
      f.Controls("lejek" & i).Visible = True
      Call ustawFormatowanieWarunkowe(i, f, rec!ID)
      rec.MoveNext
   Loop
End Sub

Sub ustawFormatowanieWarunkowe(i As Integer, f As Form, IDlejka As Integer)
   Dim tb As TextBox
   Set tb = f.Controls("lejek" & i)
   tb.Tag = IDlejka
   tb.FormatConditions.Delete
   tb.FormatConditions.Add acExpression, , "[etapWLejku]>=" & i
   tb.FormatConditions(0).BackColor = "5026082"
   tb.FormatConditions.Add acExpression, , "[etapWLejku]<" & i
   tb.FormatConditions(1).BackColor = "1643706"
   tb.FormatConditions(1).ForeColor = 16777215
End Sub

Sub ustawLejek(projekt As Variant, lejek As Variant)
   Dim sql As String
   sql = mkStr("Update projekty set etapWLejku = $1 where ID = $2", lejek, projekt)
   Call runSQL(sql)
End Sub