Option Compare Database
Option Explicit

Public Sub otworzNowe(projekt As Integer, etapWLejku As Integer)
   DoCmd.openForm "zadanie", , , , acFormAdd
   With Form_Zadanie
      .projekt.DefaultValue = projekt  ' Me.IDProjektu
      .odpowiedzialnosc.DefaultValue = getUser
      .lejek.DefaultValue = etapWLejku 'Me.etapWLejku
   End With
End Sub