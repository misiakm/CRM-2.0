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