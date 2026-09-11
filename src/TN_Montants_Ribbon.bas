Attribute VB_Name = "TN_Montants_Ribbon"
Option Explicit

' V3.4 - Word Ribbon callbacks.
' These wrappers keep the existing V3.3 public commands unchanged.

Public Sub TN33_Ruban_MontantSelectionne(ByVal control As IRibbonControl)
    TN33_ConvertirMontantSelectionne
End Sub

Public Sub TN33_Ruban_MontantsSelection(ByVal control As IRibbonControl)
    TN33_ConvertirMontantsSelection
End Sub

Public Sub TN33_Ruban_MontantsDocument(ByVal control As IRibbonControl)
    TN33_ConvertirMontantsDocument
End Sub
