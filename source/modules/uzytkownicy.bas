Option Compare Database
Option Explicit

Public w4gC4MwvGamRYmeq8BYL As Integer

Public Function getUser() As Integer
   getID = w4gC4MwvGamRYmeq8BYL
End Function

Public Sub setUser(Optional login As String)
   w4gC4MwvGamRYmeq8BYL = DLookup("id", "uzytkownicy", mkStr("login = '$1'", login))
End Sub