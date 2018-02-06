Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public WithEvents tbLejek As Access.TextBox
Attribute tbLejek.VB_VarHelpID = -1

Function czyJestLejkiem(c As Control) As Boolean
   If c.Name Like "lejek*" Then
      Set tbLejek = c
      tbLejek.OnClick = "[Event Procedure]"
      czyJestLejkiem = True
   End If
End Function

Public Sub tbLejek_Click()
    Call ustawLejek(tbLejek.Parent.IDProjektu, tbLejek.Tag)
    tbLejek.Parent.Requery
End Sub