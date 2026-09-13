Attribute VB_Name = "TN33_RegressionTests"
Option Explicit

' TN-Montants V3.4 - Golden Master regression suite.
'
' The expected values are external, frozen reference values.
' IMPORTANT: this runner never regenerates expected values from the converter.
'
' The CSV must contain:
'   id,input,expected
'
' UTF-8 / UTF-8 BOM is supported through late-bound ADODB.Stream.
' No additional reference is required.

Private Const TN33_TEST_TITLE As String = "TN-Montants Golden Master Regression"

Public Sub TN33_RunRegressionTests()
    Dim path As String
    path = TN33_SelectReferenceFile()
    If path = "" Then Exit Sub

    TN33_RunRegressionTestsFromFile path
End Sub

Public Sub TN33_RunRegressionTestsFromFile(ByVal csvPath As String)
    Dim lines As Variant
    Dim line As Variant
    Dim fields As Variant
    Dim result As String
    Dim expected As String
    Dim inputValue As String
    Dim id As String
    Dim total As Long
    Dim passed As Long
    Dim failed As Long
    Dim report As String

    On Error GoTo FatalError

    lines = TN33_ReadUtf8Lines(csvPath)
    If IsEmpty(lines) Then
        MsgBox "Impossible de lire le fichier de référence :" & vbCrLf & _
               csvPath, vbCritical, TN33_TEST_TITLE
        Exit Sub
    End If

    report = String$(72, "=") & vbCrLf
    report = report & TN33_TEST_TITLE & vbCrLf
    report = report & "Reference file: " & csvPath & vbCrLf
    report = report & String$(72, "=") & vbCrLf

    For Each line In lines
        If Trim$(CStr(line)) <> "" Then
            If Left$(Trim$(CStr(line)), 1) <> "#" Then
                fields = TN33_ParseCsvLine(CStr(line))

                If UBound(fields) >= 2 Then
                    If LCase$(Trim$(CStr(fields(0)))) <> "id" Then
                        id = Trim$(CStr(fields(0)))
                        inputValue = Trim$(CStr(fields(1)))
                        expected = CStr(fields(2))

                        total = total + 1
                        result = TN33_ConvertirPourTest(inputValue)

                        If result = expected Then
                            passed = passed + 1
                            Debug.Print "PASS"; vbTab; id; vbTab; inputValue
                        Else
                            failed = failed + 1
                            report = report & vbCrLf
                            report = report & "FAIL #" & id & "  Input: " & inputValue & vbCrLf
                            report = report & "Expected: " & expected & vbCrLf
                            report = report & "Actual:   " & result & vbCrLf
                            Debug.Print "FAIL"; vbTab; id; vbTab; inputValue
                        End If
                    End If
                End If
            End If
        End If
    Next line

    report = report & vbCrLf & String$(72, "-") & vbCrLf
    report = report & "TOTAL  : " & CStr(total) & vbCrLf
    report = report & "PASSED : " & CStr(passed) & vbCrLf
    report = report & "FAILED : " & CStr(failed) & vbCrLf
    report = report & String$(72, "-") & vbCrLf

    If failed = 0 Then
        report = report & "RESULT : PASS" & vbCrLf
        report = report & "All frozen reference results are unchanged." & vbCrLf
        MsgBox report, vbInformation, TN33_TEST_TITLE
    Else
        report = report & "RESULT : FAIL" & vbCrLf
        report = report & "DO NOT release this version until the differences are reviewed." & vbCrLf
        MsgBox report, vbCritical, TN33_TEST_TITLE
    End If

    Debug.Print report
    Exit Sub

FatalError:
    MsgBox "Regression suite error " & CStr(Err.Number) & _
           " : " & Err.Description, vbCritical, TN33_TEST_TITLE
End Sub

Private Function TN33_SelectReferenceFile() As String
    Dim dlg As FileDialog

    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Sélectionner tests\expected-results.csv"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV UTF-8", "*.csv"
        If .Show <> -1 Then Exit Function
        TN33_SelectReferenceFile = .SelectedItems(1)
    End With
End Function

Private Function TN33_ReadUtf8Lines(ByVal filePath As String) As Variant
    Dim stm As Object
    Dim txt As String

    On Error GoTo Failed

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    txt = stm.ReadText(-1)
    stm.Close

    If Len(txt) >= 1 Then
        If AscW(Left$(txt, 1)) = &HFEFF Then
            txt = Mid$(txt, 2)
        End If
    End If

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    TN33_ReadUtf8Lines = Split(txt, vbLf)
    Exit Function

Failed:
    TN33_ReadUtf8Lines = Empty
End Function

Private Function TN33_ParseCsvLine(ByVal line As String) As Variant
    Dim values() As String
    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim current As String
    Dim quoted As Boolean

    ReDim values(0 To 0)

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)

        If ch = """" Then
            If quoted And i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                current = current & """"
                i = i + 1
            Else
                quoted = Not quoted
            End If
        ElseIf ch = "," And Not quoted Then
            values(n) = current
            n = n + 1
            ReDim Preserve values(0 To n)
            current = ""
        Else
            current = current & ch
        End If
    Next i

    values(n) = current
    TN33_ParseCsvLine = values
End Function
