Option Explicit

' Exemple à placer dans le module de la feuille concernée.
' B2:B1000 = montants numériques
' D2:D1000 = montants en toutes lettres

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim zone As Range
    Dim c As Range

    Set zone = Intersect(Target, Me.Range("B2:B1000"))
    If zone Is Nothing Then Exit Sub

    On Error GoTo Sortie
    Application.EnableEvents = False

    For Each c In zone.Cells
        If Len(Trim$(CStr(c.Value))) = 0 Then
            Me.Cells(c.Row, "D").ClearContents
        Else
            Me.Cells(c.Row, "D").Value = TN_Montant(c.Value)
            Me.Cells(c.Row, "D").WrapText = True
        End If
    Next c

Sortie:
    Application.EnableEvents = True
End Sub
