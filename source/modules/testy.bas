'Modu³ s³u¿uy do testowania ró¿nych rozwi¹zañ

Option Compare Database
Option Explicit


Sub tworzTabele()
   Dim rec As Recordset, sql As String
   Set rec = CurrentDb.OpenRecordset("lejki")
   sql = "Create table tmp_abc ("
   Do Until rec.EOF
      sql = sql & "[" & rec!nazwa & "] integer, "
      rec.MoveNext
   Loop
   sql = Left(sql, Len(sql) - 2) & ")"
   Call runSQL(sql)
   
   Set rec = CurrentDb.OpenRecordset("projekty")
   
End Sub

Sub a()
   Dim c As Control
   For Each c In Forms("zadanie").Controls
      If c.ControlType = acLabel Then
         c.Caption = UCase(c.Caption)
      End If
   Next c
End Sub