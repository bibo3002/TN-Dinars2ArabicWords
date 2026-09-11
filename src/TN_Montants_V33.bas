Attribute VB_Name = "TN_Montants_V33"
Option Explicit

'====================================================================
' TN_Montants_V33 - Version 3.3
' Conversion des montants tunisiens en arabe juridique classique.
'
' Principes :
' - tous les identifiants internes portent le prefixe TN33_ ;
' - aucun mot arabe necessaire au moteur n'est ecrit en litteral :
'   il est construit par ChrW via TN33_U, pour eviter l'encodage VBE ;
' - le montant numerique original est conserve entre parentheses ;
' - le moteur reconnait les chiffres occidentaux, arabes et persans ;
' - le traitement global remplace aussi le marqueur de devise d'origine,
'   pour eviter un doublon de type "(...) d" apres conversion.
'====================================================================

Private Const TN33_VERSION As String = "3.3"

'====================================================================
' MACROS PUBLIQUES
'====================================================================

Public Sub TN33_ConvertirMontantSelectionne()
    Dim r As Range
    Dim source As String
    Dim montant As String
    Dim resultat As String

    If Selection.Type = wdSelectionIP Then
        MsgBox "Selectionnez un montant, par exemple : 2523.551 d", _
               vbExclamation, "Montant tunisien V3.3"
        Exit Sub
    End If

    Set r = Selection.Range
    source = TN33_NettoyerTexte(r.Text)
    montant = TN33_ExtraireMontant(source)

    If montant = "" Then
        MsgBox "Le montant selectionne n'est pas valide.", _
               vbExclamation, "Montant tunisien V3.3"
        Exit Sub
    End If

    resultat = TN33_ConstruireMontant(montant)
    If resultat = "" Then Exit Sub

    r.Text = resultat
    TN33_AppliquerRTL r
End Sub

Public Sub TN33_ConvertirMontantsSelection()
    Dim n As Long

    If Selection.Type = wdSelectionIP Then
        MsgBox "Selectionnez le texte contenant les montants.", _
               vbExclamation, "Montants tunisiens V3.3"
        Exit Sub
    End If

    On Error GoTo GestionErreur
    Application.ScreenUpdating = False
    n = TN33_ConvertirDansPlage(Selection.Range)
    Application.ScreenUpdating = True

    MsgBox CStr(n) & " montant(s) converti(s).", _
           vbInformation, "Montants tunisiens V3.3"
    Exit Sub

GestionErreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur " & CStr(Err.Number) & " : " & Err.Description, _
           vbCritical, "Montants tunisiens V3.3"
End Sub

Public Sub TN33_ConvertirMontantsDocument()
    Dim reponse As VbMsgBoxResult
    Dim n As Long

    reponse = MsgBox( _
        "Tous les montants associes a une devise tunisienne seront convertis." & _
        vbCrLf & "Continuer ?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Montants tunisiens V3.3")

    If reponse <> vbYes Then Exit Sub

    On Error GoTo GestionErreur
    Application.ScreenUpdating = False
    n = TN33_ConvertirDansPlage(ActiveDocument.Content)
    Application.ScreenUpdating = True

    MsgBox CStr(n) & " montant(s) converti(s).", _
           vbInformation, "Montants tunisiens V3.3"
    Exit Sub

GestionErreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur " & CStr(Err.Number) & " : " & Err.Description, _
           vbCritical, "Montants tunisiens V3.3"
End Sub

Public Sub TN33_TesterV33()
    Debug.Print String$(78, "=")
    Debug.Print "TN Montants - tests V" & TN33_VERSION
    Debug.Print String$(78, "=")

    TN33_TesterMontant "0.000"
    TN33_TesterMontant "1.000"
    TN33_TesterMontant "2.000"
    TN33_TesterMontant "6.009"
    TN33_TesterMontant "33.104"
    TN33_TesterMontant "100.000"
    TN33_TesterMontant "101.000"
    TN33_TesterMontant "200.000"
    TN33_TesterMontant "201.000"
    TN33_TesterMontant "212.000"
    TN33_TesterMontant "300.000"
    TN33_TesterMontant "999.999"
    TN33_TesterMontant "1000.000"
    TN33_TesterMontant "2000.000"
    TN33_TesterMontant "2002.000"
    TN33_TesterMontant "2523.551"
    TN33_TesterMontant "3150.000"
    TN33_TesterMontant "10004.585"
    TN33_TesterMontant "10103.341"
    TN33_TesterMontant "11429.223"
    TN33_TesterMontant "12620.405"
    TN33_TesterMontant "13804.349"
    TN33_TesterMontant "160426.591"
    TN33_TesterMontant "50845.801"
    TN33_TesterMontant "70575.300"

    MsgBox "Tests V3.3 termines. Voir la fenetre Immediate (Ctrl+G).", _
           vbInformation, "Montants tunisiens V3.3"
End Sub

'====================================================================
' CONVERSION D'UNE PLAGE WORD
'====================================================================

Private Function TN33_ConvertirDansPlage(ByVal zone As Range) As Long
    Dim r As Range
    Dim position As Long
    Dim finZone As Long
    Dim ancienLongueur As Long
    Dim montant As String
    Dim resultat As String
    Dim compteur As Long

    position = zone.Start
    finZone = zone.End

    Do While position < finZone
        Set r = TN33_ProchainMontant(zone.Document, position, finZone)
        If r Is Nothing Then Exit Do

        If TN33_EstDejaConverti(r) Then
            position = r.End
        Else
            montant = TN33_ExtraireMontant(TN33_NettoyerTexte(r.Text))

            If montant = "" Then
                position = r.End
            Else
                resultat = TN33_ConstruireMontant(montant)

                If resultat = "" Then
                    position = r.End
                Else
                    ancienLongueur = r.End - r.Start
                    r.Text = resultat
                    TN33_AppliquerRTL r

                    finZone = finZone + Len(resultat) - ancienLongueur
                    position = r.End
                    compteur = compteur + 1
                End If
            End If
        End If
    Loop

    TN33_ConvertirDansPlage = compteur
End Function

'====================================================================
' RECHERCHE D'UN MONTANT SUIVI D'UNE DEVISE
'====================================================================

Private Function TN33_ProchainMontant( _
    ByVal doc As Document, _
    ByVal debut As Long, _
    ByVal fin As Long) As Range

    Dim zone As Range
    Dim texte As String
    Dim i As Long
    Dim j As Long
    Dim finNombre As Long
    Dim posDevise As Long
    Dim longueurDevise As Long
    Dim candidat As String
    Dim montant As String

    Set zone = doc.Range(debut, fin)
    texte = zone.Text
    i = 1

    Do While i <= Len(texte)
        If TN33_EstChiffre(Mid$(texte, i, 1)) Then
            j = i

            Do While j <= Len(texte)
                If TN33_EstCaractereNombre(Mid$(texte, j, 1)) Then
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop

            finNombre = j - 1

            Do While finNombre >= i
                If TN33_EstEspace(Mid$(texte, finNombre, 1)) Then
                    finNombre = finNombre - 1
                Else
                    Exit Do
                End If
            Loop

            If finNombre >= i Then
                candidat = Mid$(texte, i, finNombre - i + 1)
                montant = TN33_NormaliserNombre(candidat)

                If montant <> "" Then
                    posDevise = finNombre + 1

                    Do While posDevise <= Len(texte)
                        If TN33_EstEspace(Mid$(texte, posDevise, 1)) Then
                            posDevise = posDevise + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    longueurDevise = TN33_LongueurDevise(texte, posDevise)

                    If longueurDevise > 0 Then
                        Set TN33_ProchainMontant = doc.Range( _
                            debut + i - 1, _
                            debut + posDevise + longueurDevise - 1)
                        Exit Function
                    End If
                End If
            End If

            If j > i Then
                i = j
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
End Function

Private Function TN33_LongueurDevise( _
    ByVal texte As String, _
    ByVal position As Long) As Long

    Dim reste As String
    Dim resteMin As String
    Dim motDinar As String
    Dim symbole As String
    Dim symboleDT As String

    If position < 1 Or position > Len(texte) Then Exit Function

    reste = Mid$(texte, position)
    resteMin = LCase$(reste)

    motDinar = TN33_W("dinar")
    symbole = TN33_Devise()
    symboleDT = symbole & "." & TN33_U("1578")

    If Left$(reste, Len(motDinar)) = motDinar Then
        TN33_LongueurDevise = Len(motDinar)
        Exit Function
    End If

    If Left$(reste, Len(symboleDT)) = symboleDT Then
        TN33_LongueurDevise = Len(symboleDT)
        Exit Function
    End If

    If Left$(resteMin, 3) = "tnd" Then
        TN33_LongueurDevise = 3
        Exit Function
    End If

    If Left$(resteMin, 2) = "dt" Then
        TN33_LongueurDevise = 2
        Exit Function
    End If

    If Left$(reste, Len(symbole)) = symbole Then
        TN33_LongueurDevise = Len(symbole)
    End If
End Function

'====================================================================
' EXTRACTION ET NORMALISATION NUMERIQUE
'====================================================================

Private Function TN33_ExtraireMontant(ByVal source As String) As String
    Dim i As Long
    Dim debutNombre As Long
    Dim finNombre As Long
    Dim ch As String

    source = TN33_NettoyerTexte(source)
    If source = "" Then Exit Function

    i = 1
    Do While i <= Len(source) And TN33_EstEspace(Mid$(source, i, 1))
        i = i + 1
    Loop

    If i > Len(source) Then Exit Function
    If Not TN33_EstChiffre(Mid$(source, i, 1)) Then Exit Function

    debutNombre = i

    Do While i <= Len(source)
        ch = Mid$(source, i, 1)
        If TN33_EstCaractereNombre(ch) Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    finNombre = i - 1
    Do While finNombre >= debutNombre
        If TN33_EstEspace(Mid$(source, finNombre, 1)) Then
            finNombre = finNombre - 1
        Else
            Exit Do
        End If
    Loop

    If finNombre < debutNombre Then Exit Function

    TN33_ExtraireMontant = TN33_NormaliserNombre( _
        Mid$(source, debutNombre, finNombre - debutNombre + 1))
End Function

Private Function TN33_NormaliserNombre(ByVal source As String) As String
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
        chiffre = TN33_ChiffreOccidental(ch)

        If chiffre <> "" Then
            brut = brut & chiffre
        ElseIf ch = "." Or ch = "," Or ch = TN33_U("1643") Then
            brut = brut & ch
        ElseIf TN33_EstEspace(ch) Or ch = TN33_U("1644") Then
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
        If ch = "." Or ch = "," Or ch = TN33_U("1643") Then
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
    entier = Replace(entier, TN33_U("1643"), "")

    If entier = "" Then Exit Function
    If Not TN33_ChaineChiffres(entier) Then Exit Function
    If decimales <> "" Then
        If Not TN33_ChaineChiffres(decimales) Then Exit Function
    End If

    entier = TN33_SupprimerZeros(entier)
    If entier = "" Then entier = "0"

    If Len(entier) > 9 Then Exit Function
    If CLng(entier) > 999999999 Then Exit Function

    decimales = decimales & "000"
    decimales = Left$(decimales, 3)

    TN33_NormaliserNombre = entier & "." & decimales
End Function

Private Function TN33_ChaineChiffres(ByVal s As String) As Boolean
    Dim i As Long

    If s = "" Then Exit Function

    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
            Exit Function
        End If
    Next i

    TN33_ChaineChiffres = True
End Function

Private Function TN33_EstCaractereNombre(ByVal ch As String) As Boolean
    If TN33_EstChiffre(ch) Then
        TN33_EstCaractereNombre = True
    ElseIf ch = "." Or ch = "," Then
        TN33_EstCaractereNombre = True
    ElseIf ch = TN33_U("1643") Or ch = TN33_U("1644") Then
        TN33_EstCaractereNombre = True
    ElseIf TN33_EstEspace(ch) Then
        TN33_EstCaractereNombre = True
    End If
End Function

Private Function TN33_EstEspace(ByVal ch As String) As Boolean
    If ch = " " Or ch = ChrW(160) Or ch = vbTab Then
        TN33_EstEspace = True
    End If
End Function

Private Function TN33_ChiffreOccidental(ByVal ch As String) As String
    Dim code As Long

    If Len(ch) = 0 Then Exit Function
    code = AscW(ch)

    If code >= 48 And code <= 57 Then
        TN33_ChiffreOccidental = Chr$(code)
        Exit Function
    End If

    If code >= &H660 And code <= &H669 Then
        TN33_ChiffreOccidental = Chr$(48 + code - &H660)
        Exit Function
    End If

    If code >= &H6F0 And code <= &H6F9 Then
        TN33_ChiffreOccidental = Chr$(48 + code - &H6F0)
    End If
End Function

Private Function TN33_EstChiffre(ByVal ch As String) As Boolean
    TN33_EstChiffre = (TN33_ChiffreOccidental(ch) <> "")
End Function

'====================================================================
' CONSTRUCTION DU MONTANT
'====================================================================

Private Function TN33_ConstruireMontant(ByVal montant As String) As String
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
        resultat = TN33_W("zero") & " " & TN33_W("dinar")
    ElseIf dinars = 0 Then
        resultat = TN33_PhraseMillimes(millimes)
    ElseIf millimes = 0 Then
        resultat = TN33_PhraseDinars(dinars)
    Else
        resultat = TN33_PhraseDinars(dinars)
        resultat = resultat & TN33_Wa() & TN33_PhraseMillimes(millimes)
    End If

    TN33_ConstruireMontant = resultat & " (" & _
                             TN33_FormaterMontant(montant) & _
                             " " & TN33_Devise() & ")"
End Function

Private Function TN33_PhraseDinars(ByVal n As Long) As String
    TN33_PhraseDinars = TN33_PhraseMonetaire( _
        n, _
        TN33_W("dinar"), _
        TN33_W("dinarOne"), _
        TN33_W("dinarTwo"), _
        TN33_W("dinarPlural"), _
        TN33_W("dinarAccusative"))
End Function

Private Function TN33_PhraseMillimes(ByVal n As Long) As String
    TN33_PhraseMillimes = TN33_PhraseMonetaire( _
        n, _
        TN33_W("millime"), _
        TN33_W("millimeOne"), _
        TN33_W("millimeTwo"), _
        TN33_W("millimePlural"), _
        TN33_W("millimeAccusative"))
End Function

Private Function TN33_PhraseMonetaire( _
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
        TN33_PhraseMonetaire = TN33_PhraseSousMilleMonetaire( _
            n, singulier, formeUn, formeDeux, pluriel, accusatif)
        Exit Function
    End If

    millions = n \ 1000000
    milliers = (n Mod 1000000) \ 1000
    reste = n Mod 1000

    If millions > 0 Then
        aUneSuite = (milliers > 0 Or reste > 0)
        resultat = TN33_EchelleMillions(millions, aUneSuite)
    End If

    If milliers > 0 Then
        If resultat <> "" Then resultat = resultat & TN33_Wa()
        aUneSuite = (reste > 0)
        resultat = resultat & TN33_EchelleMilliers(milliers, aUneSuite)
    End If

    If reste > 0 Then
        If resultat <> "" Then resultat = resultat & TN33_Wa()
        resultat = resultat & TN33_PhraseSousMilleMonetaire( _
            reste, singulier, formeUn, formeDeux, pluriel, accusatif)
    Else
        resultat = TN33_AjouterNomApresEchelle(resultat, singulier, milliers, millions)
    End If

    TN33_PhraseMonetaire = resultat
End Function

Private Function TN33_AjouterNomApresEchelle( _
    ByVal resultat As String, _
    ByVal singulier As String, _
    ByVal milliers As Long, _
    ByVal millions As Long) As String

    ' Les formes duales d'echelle exactes sont deja mises au construit :
    ' 2000 -> "alfa dinar", 2 000 000 -> "milyouna dinar".
    TN33_AjouterNomApresEchelle = resultat & " " & singulier
End Function

Private Function TN33_PhraseSousMilleMonetaire( _
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
        TN33_PhraseSousMilleMonetaire = formeUn
        Exit Function
    End If

    If n = 2 Then
        TN33_PhraseSousMilleMonetaire = formeDeux
        Exit Function
    End If

    deuxDerniers = n Mod 100
    centaines = n \ 100

    If deuxDerniers = 1 Then
        prefixe = TN33_CentainesSeules(centaines, False)
        TN33_PhraseSousMilleMonetaire = prefixe & TN33_Wa() & formeUn
        Exit Function
    End If

    If deuxDerniers = 2 Then
        prefixe = TN33_CentainesSeules(centaines, False)
        TN33_PhraseSousMilleMonetaire = prefixe & TN33_Wa() & formeDeux
        Exit Function
    End If

    If deuxDerniers >= 3 And deuxDerniers <= 10 Then
        TN33_PhraseSousMilleMonetaire = TN33_NombreSousMille(n) & _
                                        " " & pluriel
        Exit Function
    End If

    If deuxDerniers = 0 Then
        If n = 200 Then
            TN33_PhraseSousMilleMonetaire = TN33_W("hundredDualConstruct") & _
                                            " " & singulier
        Else
            TN33_PhraseSousMilleMonetaire = TN33_NombreSousMille(n) & _
                                            " " & singulier
        End If
        Exit Function
    End If

    TN33_PhraseSousMilleMonetaire = TN33_NombreSousMille(n) & _
                                    " " & accusatif
End Function

'====================================================================
' NOMBRES ET ECHELLES
'====================================================================

Private Function TN33_EchelleMilliers( _
    ByVal n As Long, _
    ByVal avecSuite As Boolean) As String

    Select Case n
        Case 1
            TN33_EchelleMilliers = TN33_W("thousand")

        Case 2
            If avecSuite Then
                TN33_EchelleMilliers = TN33_W("thousandDual")
            Else
                TN33_EchelleMilliers = TN33_W("thousandDualConstruct")
            End If

        Case 3 To 10
            TN33_EchelleMilliers = TN33_NombreSousMille(n) & _
                                    " " & TN33_W("thousandPlural")

        Case Else
            TN33_EchelleMilliers = TN33_NombreSousMille(n) & " "
            If avecSuite Then
                TN33_EchelleMilliers = TN33_EchelleMilliers & _
                                        TN33_W("thousandAccusative")
            Else
                TN33_EchelleMilliers = TN33_EchelleMilliers & _
                                        TN33_W("thousand")
            End If
    End Select
End Function

Private Function TN33_EchelleMillions( _
    ByVal n As Long, _
    ByVal avecSuite As Boolean) As String

    Select Case n
        Case 1
            TN33_EchelleMillions = TN33_W("million")

        Case 2
            If avecSuite Then
                TN33_EchelleMillions = TN33_W("millionDual")
            Else
                TN33_EchelleMillions = TN33_W("millionDualConstruct")
            End If

        Case 3 To 10
            TN33_EchelleMillions = TN33_NombreSousMille(n) & _
                                    " " & TN33_W("millionPlural")

        Case Else
            TN33_EchelleMillions = TN33_NombreSousMille(n) & " "
            If avecSuite Then
                TN33_EchelleMillions = TN33_EchelleMillions & _
                                        TN33_W("millionAccusative")
            Else
                TN33_EchelleMillions = TN33_EchelleMillions & _
                                        TN33_W("million")
            End If
    End Select
End Function

Private Function TN33_NombreSousMille(ByVal n As Long) As String
    Dim centaines As Long
    Dim reste As Long
    Dim resultat As String

    If n = 0 Then Exit Function

    If n < 100 Then
        TN33_NombreSousMille = TN33_NombreSousCent(n)
        Exit Function
    End If

    centaines = n \ 100
    reste = n Mod 100
    resultat = TN33_CentainesSeules(centaines, False)

    If reste > 0 Then
        resultat = resultat & TN33_Wa() & TN33_NombreSousCent(reste)
    End If

    TN33_NombreSousMille = resultat
End Function

Private Function TN33_CentainesSeules( _
    ByVal centaines As Long, _
    ByVal formeConstruite As Boolean) As String

    Select Case centaines
        Case 1
            TN33_CentainesSeules = TN33_W("hundred")
        Case 2
            If formeConstruite Then
                TN33_CentainesSeules = TN33_W("hundredDualConstruct")
            Else
                TN33_CentainesSeules = TN33_W("hundredDual")
            End If
        Case 3
            TN33_CentainesSeules = TN33_W("threeHundred")
        Case 4
            TN33_CentainesSeules = TN33_W("fourHundred")
        Case 5
            TN33_CentainesSeules = TN33_W("fiveHundred")
        Case 6
            TN33_CentainesSeules = TN33_W("sixHundred")
        Case 7
            TN33_CentainesSeules = TN33_W("sevenHundred")
        Case 8
            TN33_CentainesSeules = TN33_W("eightHundred")
        Case 9
            TN33_CentainesSeules = TN33_W("nineHundred")
    End Select
End Function

Private Function TN33_NombreSousCent(ByVal n As Long) As String
    Dim valeurUnite As Long
    Dim valeurDizaine As Long

    If n = 0 Then Exit Function

    If n < 10 Then
        TN33_NombreSousCent = TN33_MotUnite(n)
        Exit Function
    End If

    If n <= 19 Then
        TN33_NombreSousCent = TN33_NombreDixA19(n)
        Exit Function
    End If

    valeurDizaine = (n \ 10) * 10
    valeurUnite = n Mod 10

    If valeurUnite = 0 Then
        TN33_NombreSousCent = TN33_MotDizaine(valeurDizaine)
    Else
        TN33_NombreSousCent = TN33_MotUnite(valeurUnite) & _
                             TN33_Wa() & _
                             TN33_MotDizaine(valeurDizaine)
    End If
End Function

Private Function TN33_NombreDixA19(ByVal n As Long) As String
    Select Case n
        Case 10: TN33_NombreDixA19 = TN33_W("ten")
        Case 11: TN33_NombreDixA19 = TN33_W("eleven")
        Case 12: TN33_NombreDixA19 = TN33_W("twelve")
        Case 13: TN33_NombreDixA19 = TN33_W("thirteen")
        Case 14: TN33_NombreDixA19 = TN33_W("fourteen")
        Case 15: TN33_NombreDixA19 = TN33_W("fifteen")
        Case 16: TN33_NombreDixA19 = TN33_W("sixteen")
        Case 17: TN33_NombreDixA19 = TN33_W("seventeen")
        Case 18: TN33_NombreDixA19 = TN33_W("eighteen")
        Case 19: TN33_NombreDixA19 = TN33_W("nineteen")
    End Select
End Function

Private Function TN33_MotUnite(ByVal n As Long) As String
    Select Case n
        Case 1: TN33_MotUnite = TN33_W("one")
        Case 2: TN33_MotUnite = TN33_W("two")
        Case 3: TN33_MotUnite = TN33_W("three")
        Case 4: TN33_MotUnite = TN33_W("four")
        Case 5: TN33_MotUnite = TN33_W("five")
        Case 6: TN33_MotUnite = TN33_W("six")
        Case 7: TN33_MotUnite = TN33_W("seven")
        Case 8: TN33_MotUnite = TN33_W("eight")
        Case 9: TN33_MotUnite = TN33_W("nine")
    End Select
End Function

Private Function TN33_MotDizaine(ByVal n As Long) As String
    Select Case n
        Case 20: TN33_MotDizaine = TN33_W("twenty")
        Case 30: TN33_MotDizaine = TN33_W("thirty")
        Case 40: TN33_MotDizaine = TN33_W("forty")
        Case 50: TN33_MotDizaine = TN33_W("fifty")
        Case 60: TN33_MotDizaine = TN33_W("sixty")
        Case 70: TN33_MotDizaine = TN33_W("seventy")
        Case 80: TN33_MotDizaine = TN33_W("eighty")
        Case 90: TN33_MotDizaine = TN33_W("ninety")
    End Select
End Function

'====================================================================
' DICTIONNAIRE ARABE EN UNICODE
'====================================================================

Private Function TN33_W(ByVal nom As String) As String
    Select Case nom
        Case "zero": TN33_W = TN33_U("1589 1601 1585")
        Case "one": TN33_W = TN33_U("1608 1575 1581 1583")
        Case "two": TN33_W = TN33_U("1575 1579 1606 1575 1606")
        Case "three": TN33_W = TN33_U("1579 1604 1575 1579 1577")
        Case "four": TN33_W = TN33_U("1571 1585 1576 1593 1577")
        Case "five": TN33_W = TN33_U("1582 1605 1587 1577")
        Case "six": TN33_W = TN33_U("1587 1578 1577")
        Case "seven": TN33_W = TN33_U("1587 1576 1593 1577")
        Case "eight": TN33_W = TN33_U("1579 1605 1575 1606 1610 1577")
        Case "nine": TN33_W = TN33_U("1578 1587 1593 1577")
        Case "ten": TN33_W = TN33_U("1593 1588 1585 1577")
        Case "eleven": TN33_W = TN33_U("1571 1581 1583 32 1593 1588 1585")
        Case "twelve": TN33_W = TN33_U("1575 1579 1606 1575 32 1593 1588 1585")
        Case "thirteen": TN33_W = TN33_U("1579 1604 1575 1579 1577 32 1593 1588 1585")
        Case "fourteen": TN33_W = TN33_U("1571 1585 1576 1593 1577 32 1593 1588 1585")
        Case "fifteen": TN33_W = TN33_U("1582 1605 1587 1577 32 1593 1588 1585")
        Case "sixteen": TN33_W = TN33_U("1587 1578 1577 32 1593 1588 1585")
        Case "seventeen": TN33_W = TN33_U("1587 1576 1593 1577 32 1593 1588 1585")
        Case "eighteen": TN33_W = TN33_U("1579 1605 1575 1606 1610 1577 32 1593 1588 1585")
        Case "nineteen": TN33_W = TN33_U("1578 1587 1593 1577 32 1593 1588 1585")
        Case "twenty": TN33_W = TN33_U("1593 1588 1585 1608 1606")
        Case "thirty": TN33_W = TN33_U("1579 1604 1575 1579 1608 1606")
        Case "forty": TN33_W = TN33_U("1571 1585 1576 1593 1608 1606")
        Case "fifty": TN33_W = TN33_U("1582 1605 1587 1608 1606")
        Case "sixty": TN33_W = TN33_U("1587 1578 1608 1606")
        Case "seventy": TN33_W = TN33_U("1587 1576 1593 1608 1606")
        Case "eighty": TN33_W = TN33_U("1579 1605 1575 1606 1608 1606")
        Case "ninety": TN33_W = TN33_U("1578 1587 1593 1608 1606")

        Case "hundred": TN33_W = TN33_U("1605 1575 1574 1577")
        Case "hundredDual": TN33_W = TN33_U("1605 1575 1574 1578 1575 1606")
        Case "hundredDualConstruct": TN33_W = TN33_U("1605 1575 1574 1578 1575")
        Case "threeHundred": TN33_W = TN33_U("1579 1604 1575 1579 1605 1575 1574 1577")
        Case "fourHundred": TN33_W = TN33_U("1571 1585 1576 1593 1605 1575 1574 1577")
        Case "fiveHundred": TN33_W = TN33_U("1582 1605 1587 1605 1575 1574 1577")
        Case "sixHundred": TN33_W = TN33_U("1587 1578 1605 1575 1574 1577")
        Case "sevenHundred": TN33_W = TN33_U("1587 1576 1593 1605 1575 1574 1577")
        Case "eightHundred": TN33_W = TN33_U("1579 1605 1575 1606 1605 1575 1574 1577")
        Case "nineHundred": TN33_W = TN33_U("1578 1587 1593 1605 1575 1574 1577")

        Case "thousand": TN33_W = TN33_U("1571 1604 1601")
        Case "thousandDual": TN33_W = TN33_U("1571 1604 1601 1575 1606")
        Case "thousandDualConstruct": TN33_W = TN33_U("1571 1604 1601 1575")
        Case "thousandPlural": TN33_W = TN33_U("1570 1604 1575 1601")
        Case "thousandAccusative": TN33_W = TN33_U("1571 1604 1601 1575")

        Case "million": TN33_W = TN33_U("1605 1604 1610 1608 1606")
        Case "millionDual": TN33_W = TN33_U("1605 1604 1610 1608 1606 1575 1606")
        Case "millionDualConstruct": TN33_W = TN33_U("1605 1604 1610 1608 1606 1575")
        Case "millionPlural": TN33_W = TN33_U("1605 1604 1575 1610 1610 1606")
        Case "millionAccusative": TN33_W = TN33_U("1605 1604 1610 1608 1606 1575")

        Case "dinar": TN33_W = TN33_U("1583 1610 1606 1575 1585")
        Case "dinarOne": TN33_W = TN33_U("1583 1610 1606 1575 1585 32 1608 1575 1581 1583")
        Case "dinarTwo": TN33_W = TN33_U("1583 1610 1606 1575 1585 1575 1606")
        Case "dinarPlural": TN33_W = TN33_U("1583 1606 1575 1606 1610 1585")
        Case "dinarAccusative": TN33_W = TN33_U("1583 1610 1606 1575 1585 1575")

        Case "millime": TN33_W = TN33_U("1605 1604 1610 1605")
        Case "millimeOne": TN33_W = TN33_U("1605 1604 1610 1605 32 1608 1575 1581 1583")
        Case "millimeTwo": TN33_W = TN33_U("1605 1604 1610 1605 1575 1606")
        Case "millimePlural": TN33_W = TN33_U("1605 1604 1610 1605 1575 1578")
        Case "millimeAccusative": TN33_W = TN33_U("1605 1604 1610 1605 1575")
    End Select
End Function

Private Function TN33_U(ByVal codes As String) As String
    Dim elements() As String
    Dim i As Long
    Dim resultat As String

    elements = Split(codes, " ")

    For i = LBound(elements) To UBound(elements)
        resultat = resultat & ChrW(CLng(elements(i)))
    Next i

    TN33_U = resultat
End Function

Private Function TN33_Wa() As String
    TN33_Wa = " " & ChrW(&H648)
End Function

Private Function TN33_Devise() As String
    TN33_Devise = ChrW(&H62F)
End Function

'====================================================================
' UTILITAIRES WORD ET FORMATAGE
'====================================================================

Private Function TN33_FormaterMontant(ByVal montant As String) As String
    Dim p As Long
    Dim entier As String
    Dim decimales As String

    p = InStr(1, montant, ".", vbBinaryCompare)

    If p = 0 Then
        TN33_FormaterMontant = montant & ".000"
        Exit Function
    End If

    entier = Left$(montant, p - 1)
    decimales = Mid$(montant, p + 1) & "000"
    decimales = Left$(decimales, 3)

    TN33_FormaterMontant = entier & "." & decimales
End Function

Private Function TN33_SupprimerZeros(ByVal s As String) As String
    Do While Len(s) > 1 And Left$(s, 1) = "0"
        s = Mid$(s, 2)
    Loop

    TN33_SupprimerZeros = s
End Function

Private Function TN33_NettoyerTexte(ByVal s As String) As String
    s = Replace(s, Chr$(7), "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    TN33_NettoyerTexte = Trim$(s)
End Function

Private Function TN33_EstDejaConverti(ByVal r As Range) As Boolean
    Dim avant As Range
    Dim apres As Range
    Dim chAvant As String
    Dim chApres As String
    Dim p As Long

    On Error GoTo Fin

    p = r.Start - 1
    Do While p >= 0
        Set avant = r.Document.Range(p, p + 1)
        chAvant = avant.Text
        If TN33_EstEspace(chAvant) Then
            p = p - 1
        Else
            Exit Do
        End If
    Loop

    If chAvant <> "(" Then Exit Function

    p = r.End
    Do While p < r.Document.Content.End
        Set apres = r.Document.Range(p, p + 1)
        chApres = apres.Text
        If TN33_EstEspace(chApres) Then
            p = p + 1
        Else
            Exit Do
        End If
    Loop

    If chApres = ")" Then TN33_EstDejaConverti = True

Fin:
End Function

Private Sub TN33_AppliquerRTL(ByVal r As Range)
    On Error Resume Next
    r.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
    On Error GoTo 0
End Sub

Private Sub TN33_TesterMontant(ByVal montant As String)
    Debug.Print montant & " -> " & TN33_ConstruireMontant(montant)
End Sub
