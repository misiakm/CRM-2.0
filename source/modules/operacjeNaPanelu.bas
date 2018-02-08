Option Compare Database
Option Explicit

Sub pokazUkryjMalyPrawy(ukryj As Boolean)
   Form_panel.btnSzanseSprzedazy.SetFocus
   Form_panel.f_malyPrawy.Visible = Not ukryj
End Sub

Sub pokazUkryjMalyLewy(ukryj As Boolean)
   Form_panel.btnSzanseSprzedazy.SetFocus
   Form_panel.f_malyLewy.Visible = Not ukryj
End Sub

Sub pokazUkryjDuzy(ukryj As Boolean)
   Form_panel.btnSzanseSprzedazy.SetFocus
   Form_panel.f_duzy.Visible = Not ukryj
End Sub

Sub ustawMale(formLewy As String, formPrawy As String, Optional filtrLewy As String, Optional filtrPrawy As String)
   Call ustawPodformularz(Form_panel.f_malyLewy, formLewy, filtrLewy)
   Call ustawPodformularz(Form_panel.f_malyPrawy, formPrawy, filtrPrawy)
   Call pokazUkryjDuzy(True)
   Call pokazUkryjMalyLewy(False)
   Call pokazUkryjMalyPrawy(False)
End Sub

Sub ustawDuzy(formularz As String, Optional filtr As String)
   Call ustawPodformularz(Form_panel.f_duzy, formularz, filtr)
   Call pokazUkryjDuzy(False)
   Call pokazUkryjMalyLewy(True)
   Call pokazUkryjMalyPrawy(True)
End Sub

Sub ustawPodformularz(f As Object, obiekt As String, filtr As String)
   With f
      If czyJestraportem(obiekt) Then
         .SourceObject = "Raport." & obiekt
         .Report.FilterOn = False
         .Report.Filter = filtr
         .Report.FilterOn = filtr <> ""
      Else
         .SourceObject = obiekt
         .Form.FilterOn = False
         .Form.Filter = filtr
         .Form.FilterOn = filtr <> ""
      End If
   End With
End Sub

Function czyJestraportem(obiekt As String) As Boolean
   On Error GoTo blad
   Echo False
   DoCmd.openReport obiekt, acViewReport
   czyJestraportem = True
   DoCmd.Close acReport, obiekt
blad:
   Echo True
End Function