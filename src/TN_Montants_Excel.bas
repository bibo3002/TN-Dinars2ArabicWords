Attribute VB_Name = "TN_Montants_Excel"
Option Explicit

' TN-Montants V3.4 - Excel engine.
' This module contains the non-Word conversion engine derived from V3.3.
' It has no Word object dependencies.
'
' Worksheet formula:
'     =TN_Montant(B2)
'
Public Function TN_Montant(ByVal valeur As Variant) As String
    Dim source As String
    Dim montant As String

    If IsError(valeur) Then Exit Function
    If IsEmpty(valeur) Then Exit Function

    source = Trim$(CStr(valeur))
    If source = "" Then Exit Function

    montant = TN34X_ExtraireMontant(source)
    If montant = "" Then
        ' Excel numeric cells can arrive as 2523.551 depending on locale.
        ' Try the normalized text directly as a fallback.
        montant = TN34X_NormaliserNombre(source)
    End If

    If montant = "" Then Exit Function
    TN_Montant = TN34X_ConstruireMontant(montant)
End Function

Public Function TN_MontantNombre(ByVal valeur As Variant) As String
    TN_MontantNombre = TN_Montant(valeur)
End Function

Private Function TN34X_ExtraireMontant(ByVal source As String) As String
    Dim i As Long
    Dim debutNombre As Long
    Dim finNombre As Long
    Dim ch As String

    source = TN34X_NettoyerTexte(source)
    If source = "" Then Exit Function

    i = 1
    Do While i <= Len(source) And TN34X_EstEspace(Mid$(source, i, 1))
        i = i + 1
    Loop

    If i > Len(source) Then Exit Function
    If Not TN34X_EstChiffre(Mid$(source, i, 1)) Then Exit Function

    debutNombre = i

    Do While i <= Len(source)
        ch = Mid$(source, i, 1)
        If TN34X_EstCaractereNombre(ch) Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    finNombre = i - 1
    Do While finNombre >= debutNombre
        If TN34X_EstEspace(Mid$(source, finNombre, 1)) Then
            finNombre = finNombre - 1
        Else
            Exit Do
        End If
    Loop

    If finNombre < debutNombre Then Exit Function

    TN34X_ExtraireMontant = TN34X_NormaliserNombre( _
        Mid$(source, debutNombre, finNombre - debutNombre + 1))
End Function

Private Function TN34X_NormaliserNombre(ByVal source As String) As String
    Dim i As Long
    Dim ch As String
    Dim chiffre As String
    Dim brut As String
    Dim posDecimal As Long
    Dim nbApres As Long
    Dim entier As String
    Dim decimales As String
    Dim separateur As String

    source = Trim$(source)
    If source = "" Then Exit Function

    ' Conserver uniquement chiffres et separateurs numeriques autorises.
    For i = 1 To Len(source)
        ch = Mid$(source, i, 1)
        chiffre = TN34X_ChiffreOccidental(ch)

        If chiffre <> "" Then
            brut = brut & chiffre
        ElseIf ch = "." Or ch = "," Or ch = TN34X_U("1643") Then
            brut = brut & ch
        ElseIf TN34X_EstEspace(ch) Or ch = TN34X_U("1644") Then
            ' espace ou separateur de milliers arabe : ignore
        Else
            Exit Function
        End If
    Next i

    If brut = "" Then Exit Function

    ' Le dernier . , ou separateur decimal arabe est considere decimal
    ' uniquement s'il est suivi de 1 a 3 chiffres.
    For i = Len(brut) To 1 Step -1
        ch = Mid$(brut, i, 1)
        If ch = "." Or ch = "," Or ch = TN34X_U("1643") Then
            nbApres = Len(brut) - i
            If nbApres >= 1 And nbApres <= 3 Then
                posDecimal = i
                separateur = ch
            End If
            Exit For
        End If
    Next i

    If posDecimal > 0 Then
        entier = Left$(brut, posDecimal - 1)
        decimales = Mid$(brut, posDecimal + 1)
    Else
        entier = brut
        decimales = ""
    End If

    ' Les separateurs restants dans la partie entiere sont des milliers.
    entier = Replace(entier, ".", "")
    entier = Replace(entier, ",", "")
    entier = Replace(entier, TN34X_U("1643"), "")

    If entier = "" Then Exit Function
    If Not TN34X_ChaineChiffres(entier) Then Exit Function
    If decimales <> "" Then
        If Not TN34X_ChaineChiffres(decimales) Then Exit Function
    End If

    entier = TN34X_SupprimerZeros(entier)
    If entier = "" Then entier = "0"

    If Len(entier) > 9 Then Exit Function
    If CLng(entier) > 999999999 Then Exit Function

    decimales = decimales & "000"
    decimales = Left$(decimales, 3)

    TN34X_NormaliserNombre = entier & "." & decimales
End Function

Private Function TN34X_ChaineChiffres(ByVal s As String) As Boolean
    Dim i As Long

    If s = "" Then Exit Function

    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
            Exit Function
        End If
    Next i

    TN34X_ChaineChiffres = True
End Function

Private Function TN34X_EstCaractereNombre(ByVal ch As String) As Boolean
    If TN34X_EstChiffre(ch) Then
        TN34X_EstCaractereNombre = True
    ElseIf ch = "." Or ch = "," Then
        TN34X_EstCaractereNombre = True
    ElseIf ch = TN34X_U("1643") Or ch = TN34X_U("1644") Then
        TN34X_EstCaractereNombre = True
    ElseIf TN34X_EstEspace(ch) Then
        TN34X_EstCaractereNombre = True
    End If
End Function

Private Function TN34X_EstEspace(ByVal ch As String) As Boolean
    If ch = " " Or ch = ChrW(160) Or ch = vbTab Then
        TN34X_EstEspace = True
    End If
End Function

Private Function TN34X_ChiffreOccidental(ByVal ch As String) As String
    Dim code As Long

    If Len(ch) = 0 Then Exit Function
    code = AscW(ch)

    If code >= 48 And code <= 57 Then
        TN34X_ChiffreOccidental = Chr$(code)
        Exit Function
    End If

    If code >= &H660 And code <= &H669 Then
        TN34X_ChiffreOccidental = Chr$(48 + code - &H660)
        Exit Function
    End If

    If code >= &H6F0 And code <= &H6F9 Then
        TN34X_ChiffreOccidental = Chr$(48 + code - &H6F0)
    End If
End Function

Private Function TN34X_EstChiffre(ByVal ch As String) As Boolean
    TN34X_EstChiffre = (TN34X_ChiffreOccidental(ch) <> "")
End Function

'====================================================================
' CONSTRUCTION DU MONTANT
'====================================================================

Private Function TN34X_ConstruireMontant(ByVal montant As String) As String
    Dim dinarsTexte As String
    Dim millimesTexte As String
    Dim dinars As Long
    Dim millimes As Long
    Dim p As Long
    Dim resultat As String

    p = InStr(1, montant, ".", vbBinaryCompare)
    If p = 0 Then Exit Function

    dinarsTexte = Left$(montant, p - 1)
    millimesTexte = Mid$(montant, p + 1)

    If dinarsTexte = "" Or millimesTexte = "" Then Exit Function
    If Len(millimesTexte) <> 3 Then Exit Function

    dinars = CLng(dinarsTexte)
    millimes = CLng(millimesTexte)

    If dinars = 0 And millimes = 0 Then
        resultat = TN34X_W("zero") & " " & TN34X_W("dinar")
    ElseIf dinars = 0 Then
        resultat = TN34X_PhraseMillimes(millimes)
    ElseIf millimes = 0 Then
        resultat = TN34X_PhraseDinars(dinars)
    Else
        resultat = TN34X_PhraseDinars(dinars)
        resultat = resultat & TN34X_Wa() & TN34X_PhraseMillimes(millimes)
    End If

    TN34X_ConstruireMontant = resultat & " (" & _
                             TN34X_FormaterMontant(montant) & _
                             " " & TN34X_Devise() & ")"
End Function

Private Function TN34X_PhraseDinars(ByVal n As Long) As String
    TN34X_PhraseDinars = TN34X_PhraseMonetaire( _
        n, _
        TN34X_W("dinar"), _
        TN34X_W("dinarOne"), _
        TN34X_W("dinarTwo"), _
        TN34X_W("dinarPlural"), _
        TN34X_W("dinarAccusative"))
End Function

Private Function TN34X_PhraseMillimes(ByVal n As Long) As String
    TN34X_PhraseMillimes = TN34X_PhraseMonetaire( _
        n, _
        TN34X_W("millime"), _
        TN34X_W("millimeOne"), _
        TN34X_W("millimeTwo"), _
        TN34X_W("millimePlural"), _
        TN34X_W("millimeAccusative"))
End Function

Private Function TN34X_PhraseMonetaire( _
    ByVal n As Long, _
    ByVal singulier As String, _
    ByVal formeUn As String, _
    ByVal formeDeux As String, _
    ByVal pluriel As String, _
    ByVal accusatif As String) As String

    Dim millions As Long
    Dim milliers As Long
    Dim reste As Long
    Dim resultat As String
    Dim aUneSuite As Boolean

    If n = 0 Then Exit Function

    If n < 1000 Then
        TN34X_PhraseMonetaire = TN34X_PhraseSousMilleMonetaire( _
            n, singulier, formeUn, formeDeux, pluriel, accusatif)
        Exit Function
    End If

    millions = n \ 1000000
    milliers = (n Mod 1000000) \ 1000
    reste = n Mod 1000

    If millions > 0 Then
        aUneSuite = (milliers > 0 Or reste > 0)
        resultat = TN34X_EchelleMillions(millions, aUneSuite)
    End If

    If milliers > 0 Then
        If resultat <> "" Then resultat = resultat & TN34X_Wa()
        aUneSuite = (reste > 0)
        resultat = resultat & TN34X_EchelleMilliers(milliers, aUneSuite)
    End If

    If reste > 0 Then
        If resultat <> "" Then resultat = resultat & TN34X_Wa()
        resultat = resultat & TN34X_PhraseSousMilleMonetaire( _
            reste, singulier, formeUn, formeDeux, pluriel, accusatif)
    Else
        resultat = TN34X_AjouterNomApresEchelle(resultat, singulier, milliers, millions)
    End If

    TN34X_PhraseMonetaire = resultat
End Function

Private Function TN34X_AjouterNomApresEchelle( _
    ByVal resultat As String, _
    ByVal singulier As String, _
    ByVal milliers As Long, _
    ByVal millions As Long) As String

    ' Les formes duales d'echelle exactes sont deja mises au construit :
    ' 2000 -> "alfa dinar", 2 000 000 -> "milyouna dinar".
    TN34X_AjouterNomApresEchelle = resultat & " " & singulier
End Function

Private Function TN34X_PhraseSousMilleMonetaire( _
    ByVal n As Long, _
    ByVal singulier As String, _
    ByVal formeUn As String, _
    ByVal formeDeux As String, _
    ByVal pluriel As String, _
    ByVal accusatif As String) As String

    Dim deuxDerniers As Long
    Dim centaines As Long
    Dim prefixe As String

    If n = 1 Then
        TN34X_PhraseSousMilleMonetaire = formeUn
        Exit Function
    End If

    If n = 2 Then
        TN34X_PhraseSousMilleMonetaire = formeDeux
        Exit Function
    End If

    deuxDerniers = n Mod 100
    centaines = n \ 100

    If deuxDerniers = 1 Then
        prefixe = TN34X_CentainesSeules(centaines, False)
        TN34X_PhraseSousMilleMonetaire = prefixe & TN34X_Wa() & formeUn
        Exit Function
    End If

    If deuxDerniers = 2 Then
        prefixe = TN34X_CentainesSeules(centaines, False)
        TN34X_PhraseSousMilleMonetaire = prefixe & TN34X_Wa() & formeDeux
        Exit Function
    End If

    If deuxDerniers >= 3 And deuxDerniers <= 10 Then
        TN34X_PhraseSousMilleMonetaire = TN34X_NombreSousMille(n) & _
                                        " " & pluriel
        Exit Function
    End If

    If deuxDerniers = 0 Then
        If n = 200 Then
            TN34X_PhraseSousMilleMonetaire = TN34X_W("hundredDualConstruct") & _
                                            " " & singulier
        Else
            TN34X_PhraseSousMilleMonetaire = TN34X_NombreSousMille(n) & _
                                            " " & singulier
        End If
        Exit Function
    End If

    TN34X_PhraseSousMilleMonetaire = TN34X_NombreSousMille(n) & _
                                    " " & accusatif
End Function

'====================================================================
' NOMBRES ET ECHELLES
'====================================================================

Private Function TN34X_EchelleMilliers( _
    ByVal n As Long, _
    ByVal avecSuite As Boolean) As String

    Select Case n
        Case 1
            TN34X_EchelleMilliers = TN34X_W("thousand")

        Case 2
            If avecSuite Then
                TN34X_EchelleMilliers = TN34X_W("thousandDual")
            Else
                TN34X_EchelleMilliers = TN34X_W("thousandDualConstruct")
            End If

        Case 3 To 10
            TN34X_EchelleMilliers = TN34X_NombreSousMille(n) & _
                                    " " & TN34X_W("thousandPlural")

        Case Else
            TN34X_EchelleMilliers = TN34X_NombreSousMille(n) & " "
            If avecSuite Then
                TN34X_EchelleMilliers = TN34X_EchelleMilliers & _
                                        TN34X_W("thousandAccusative")
            Else
                TN34X_EchelleMilliers = TN34X_EchelleMilliers & _
                                        TN34X_W("thousand")
            End If
    End Select
End Function

Private Function TN34X_EchelleMillions( _
    ByVal n As Long, _
    ByVal avecSuite As Boolean) As String

    Select Case n
        Case 1
            TN34X_EchelleMillions = TN34X_W("million")

        Case 2
            If avecSuite Then
                TN34X_EchelleMillions = TN34X_W("millionDual")
            Else
                TN34X_EchelleMillions = TN34X_W("millionDualConstruct")
            End If

        Case 3 To 10
            TN34X_EchelleMillions = TN34X_NombreSousMille(n) & _
                                    " " & TN34X_W("millionPlural")

        Case Else
            TN34X_EchelleMillions = TN34X_NombreSousMille(n) & " "
            If avecSuite Then
                TN34X_EchelleMillions = TN34X_EchelleMillions & _
                                        TN34X_W("millionAccusative")
            Else
                TN34X_EchelleMillions = TN34X_EchelleMillions & _
                                        TN34X_W("million")
            End If
    End Select
End Function

Private Function TN34X_NombreSousMille(ByVal n As Long) As String
    Dim centaines As Long
    Dim reste As Long
    Dim resultat As String

    If n = 0 Then Exit Function

    If n < 100 Then
        TN34X_NombreSousMille = TN34X_NombreSousCent(n)
        Exit Function
    End If

    centaines = n \ 100
    reste = n Mod 100
    resultat = TN34X_CentainesSeules(centaines, False)

    If reste > 0 Then
        resultat = resultat & TN34X_Wa() & TN34X_NombreSousCent(reste)
    End If

    TN34X_NombreSousMille = resultat
End Function

Private Function TN34X_CentainesSeules( _
    ByVal centaines As Long, _
    ByVal formeConstruite As Boolean) As String

    Select Case centaines
        Case 1
            TN34X_CentainesSeules = TN34X_W("hundred")
        Case 2
            If formeConstruite Then
                TN34X_CentainesSeules = TN34X_W("hundredDualConstruct")
            Else
                TN34X_CentainesSeules = TN34X_W("hundredDual")
            End If
        Case 3
            TN34X_CentainesSeules = TN34X_W("threeHundred")
        Case 4
            TN34X_CentainesSeules = TN34X_W("fourHundred")
        Case 5
            TN34X_CentainesSeules = TN34X_W("fiveHundred")
        Case 6
            TN34X_CentainesSeules = TN34X_W("sixHundred")
        Case 7
            TN34X_CentainesSeules = TN34X_W("sevenHundred")
        Case 8
            TN34X_CentainesSeules = TN34X_W("eightHundred")
        Case 9
            TN34X_CentainesSeules = TN34X_W("nineHundred")
    End Select
End Function

Private Function TN34X_NombreSousCent(ByVal n As Long) As String
    Dim valeurUnite As Long
    Dim valeurDizaine As Long

    If n = 0 Then Exit Function

    If n < 10 Then
        TN34X_NombreSousCent = TN34X_MotUnite(n)
        Exit Function
    End If

    If n <= 19 Then
        TN34X_NombreSousCent = TN34X_NombreDixA19(n)
        Exit Function
    End If

    valeurDizaine = (n \ 10) * 10
    valeurUnite = n Mod 10

    If valeurUnite = 0 Then
        TN34X_NombreSousCent = TN34X_MotDizaine(valeurDizaine)
    Else
        TN34X_NombreSousCent = TN34X_MotUnite(valeurUnite) & _
                             TN34X_Wa() & _
                             TN34X_MotDizaine(valeurDizaine)
    End If
End Function

Private Function TN34X_NombreDixA19(ByVal n As Long) As String
    Select Case n
        Case 10: TN34X_NombreDixA19 = TN34X_W("ten")
        Case 11: TN34X_NombreDixA19 = TN34X_W("eleven")
        Case 12: TN34X_NombreDixA19 = TN34X_W("twelve")
        Case 13: TN34X_NombreDixA19 = TN34X_W("thirteen")
        Case 14: TN34X_NombreDixA19 = TN34X_W("fourteen")
        Case 15: TN34X_NombreDixA19 = TN34X_W("fifteen")
        Case 16: TN34X_NombreDixA19 = TN34X_W("sixteen")
        Case 17: TN34X_NombreDixA19 = TN34X_W("seventeen")
        Case 18: TN34X_NombreDixA19 = TN34X_W("eighteen")
        Case 19: TN34X_NombreDixA19 = TN34X_W("nineteen")
    End Select
End Function

Private Function TN34X_MotUnite(ByVal n As Long) As String
    Select Case n
        Case 1: TN34X_MotUnite = TN34X_W("one")
        Case 2: TN34X_MotUnite = TN34X_W("two")
        Case 3: TN34X_MotUnite = TN34X_W("three")
        Case 4: TN34X_MotUnite = TN34X_W("four")
        Case 5: TN34X_MotUnite = TN34X_W("five")
        Case 6: TN34X_MotUnite = TN34X_W("six")
        Case 7: TN34X_MotUnite = TN34X_W("seven")
        Case 8: TN34X_MotUnite = TN34X_W("eight")
        Case 9: TN34X_MotUnite = TN34X_W("nine")
    End Select
End Function

Private Function TN34X_MotDizaine(ByVal n As Long) As String
    Select Case n
        Case 20: TN34X_MotDizaine = TN34X_W("twenty")
        Case 30: TN34X_MotDizaine = TN34X_W("thirty")
        Case 40: TN34X_MotDizaine = TN34X_W("forty")
        Case 50: TN34X_MotDizaine = TN34X_W("fifty")
        Case 60: TN34X_MotDizaine = TN34X_W("sixty")
        Case 70: TN34X_MotDizaine = TN34X_W("seventy")
        Case 80: TN34X_MotDizaine = TN34X_W("eighty")
        Case 90: TN34X_MotDizaine = TN34X_W("ninety")
    End Select
End Function

'====================================================================
' DICTIONNAIRE ARABE EN UNICODE
'====================================================================

Private Function TN34X_W(ByVal nom As String) As String
    Select Case nom
        Case "zero": TN34X_W = TN34X_U("1589 1601 1585")
        Case "one": TN34X_W = TN34X_U("1608 1575 1581 1583")
        Case "two": TN34X_W = TN34X_U("1575 1579 1606 1575 1606")
        Case "three": TN34X_W = TN34X_U("1579 1604 1575 1579 1577")
        Case "four": TN34X_W = TN34X_U("1571 1585 1576 1593 1577")
        Case "five": TN34X_W = TN34X_U("1582 1605 1587 1577")
        Case "six": TN34X_W = TN34X_U("1587 1578 1577")
        Case "seven": TN34X_W = TN34X_U("1587 1576 1593 1577")
        Case "eight": TN34X_W = TN34X_U("1579 1605 1575 1606 1610 1577")
        Case "nine": TN34X_W = TN34X_U("1578 1587 1593 1577")
        Case "ten": TN34X_W = TN34X_U("1593 1588 1585 1577")
        Case "eleven": TN34X_W = TN34X_U("1571 1581 1583 32 1593 1588 1585")
        Case "twelve": TN34X_W = TN34X_U("1575 1579 1606 1575 32 1593 1588 1585")
        Case "thirteen": TN34X_W = TN34X_U("1579 1604 1575 1579 1577 32 1593 1588 1585")
        Case "fourteen": TN34X_W = TN34X_U("1571 1585 1576 1593 1577 32 1593 1588 1585")
        Case "fifteen": TN34X_W = TN34X_U("1582 1605 1587 1577 32 1593 1588 1585")
        Case "sixteen": TN34X_W = TN34X_U("1587 1578 1577 32 1593 1588 1585")
        Case "seventeen": TN34X_W = TN34X_U("1587 1576 1593 1577 32 1593 1588 1585")
        Case "eighteen": TN34X_W = TN34X_U("1579 1605 1575 1606 1610 1577 32 1593 1588 1585")
        Case "nineteen": TN34X_W = TN34X_U("1578 1587 1593 1577 32 1593 1588 1585")
        Case "twenty": TN34X_W = TN34X_U("1593 1588 1585 1608 1606")
        Case "thirty": TN34X_W = TN34X_U("1579 1604 1575 1579 1608 1606")
        Case "forty": TN34X_W = TN34X_U("1571 1585 1576 1593 1608 1606")
        Case "fifty": TN34X_W = TN34X_U("1582 1605 1587 1608 1606")
        Case "sixty": TN34X_W = TN34X_U("1587 1578 1608 1606")
        Case "seventy": TN34X_W = TN34X_U("1587 1576 1593 1608 1606")
        Case "eighty": TN34X_W = TN34X_U("1579 1605 1575 1606 1608 1606")
        Case "ninety": TN34X_W = TN34X_U("1578 1587 1593 1608 1606")

        Case "hundred": TN34X_W = TN34X_U("1605 1575 1574 1577")
        Case "hundredDual": TN34X_W = TN34X_U("1605 1575 1574 1578 1575 1606")
        Case "hundredDualConstruct": TN34X_W = TN34X_U("1605 1575 1574 1578 1575")
        Case "threeHundred": TN34X_W = TN34X_U("1579 1604 1575 1579 1605 1575 1574 1577")
        Case "fourHundred": TN34X_W = TN34X_U("1571 1585 1576 1593 1605 1575 1574 1577")
        Case "fiveHundred": TN34X_W = TN34X_U("1582 1605 1587 1605 1575 1574 1577")
        Case "sixHundred": TN34X_W = TN34X_U("1587 1578 1605 1575 1574 1577")
        Case "sevenHundred": TN34X_W = TN34X_U("1587 1576 1593 1605 1575 1574 1577")
        Case "eightHundred": TN34X_W = TN34X_U("1579 1605 1575 1606 1605 1575 1574 1577")
        Case "nineHundred": TN34X_W = TN34X_U("1578 1587 1593 1605 1575 1574 1577")

        Case "thousand": TN34X_W = TN34X_U("1571 1604 1601")
        Case "thousandDual": TN34X_W = TN34X_U("1571 1604 1601 1575 1606")
        Case "thousandDualConstruct": TN34X_W = TN34X_U("1571 1604 1601 1575")
        Case "thousandPlural": TN34X_W = TN34X_U("1570 1604 1575 1601")
        Case "thousandAccusative": TN34X_W = TN34X_U("1571 1604 1601 1575")

        Case "million": TN34X_W = TN34X_U("1605 1604 1610 1608 1606")
        Case "millionDual": TN34X_W = TN34X_U("1605 1604 1610 1608 1606 1575 1606")
        Case "millionDualConstruct": TN34X_W = TN34X_U("1605 1604 1610 1608 1606 1575")
        Case "millionPlural": TN34X_W = TN34X_U("1605 1604 1575 1610 1610 1606")
        Case "millionAccusative": TN34X_W = TN34X_U("1605 1604 1610 1608 1606 1575")

        Case "dinar": TN34X_W = TN34X_U("1583 1610 1606 1575 1585")
        Case "dinarOne": TN34X_W = TN34X_U("1583 1610 1606 1575 1585 32 1608 1575 1581 1583")
        Case "dinarTwo": TN34X_W = TN34X_U("1583 1610 1606 1575 1585 1575 1606")
        Case "dinarPlural": TN34X_W = TN34X_U("1583 1606 1575 1606 1610 1585")
        Case "dinarAccusative": TN34X_W = TN34X_U("1583 1610 1606 1575 1585 1575")

        Case "millime": TN34X_W = TN34X_U("1605 1604 1610 1605")
        Case "millimeOne": TN34X_W = TN34X_U("1605 1604 1610 1605 32 1608 1575 1581 1583")
        Case "millimeTwo": TN34X_W = TN34X_U("1605 1604 1610 1605 1575 1606")
        Case "millimePlural": TN34X_W = TN34X_U("1605 1604 1610 1605 1575 1578")
        Case "millimeAccusative": TN34X_W = TN34X_U("1605 1604 1610 1605 1575")
    End Select
End Function

Private Function TN34X_U(ByVal codes As String) As String
    Dim elements() As String
    Dim i As Long
    Dim resultat As String

    elements = Split(codes, " ")

    For i = LBound(elements) To UBound(elements)
        resultat = resultat & ChrW(CLng(elements(i)))
    Next i

    TN34X_U = resultat
End Function

Private Function TN34X_Wa() As String
    TN34X_Wa = " " & ChrW(&H648)
End Function

Private Function TN34X_Devise() As String
    TN34X_Devise = ChrW(&H62F)
End Function

'====================================================================
' UTILITAIRES WORD ET FORMATAGE
'====================================================================

Private Function TN34X_FormaterMontant(ByVal montant As String) As String
    Dim p As Long
    Dim entier As String
    Dim decimales As String

    p = InStr(1, montant, ".", vbBinaryCompare)

    If p = 0 Then
        TN34X_FormaterMontant = montant & ".000"
        Exit Function
    End If

    entier = Left$(montant, p - 1)
    decimales = Mid$(montant, p + 1) & "000"
    decimales = Left$(decimales, 3)

    TN34X_FormaterMontant = entier & "." & decimales
End Function

Private Function TN34X_SupprimerZeros(ByVal s As String) As String
