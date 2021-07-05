Attribute VB_Name = "Module1"

' *******************************************************************************
'                               NotaComp
'
'   Outil Excel d�di� a la notation par comp�tences en milieu scolaire
'
'   Classeur vierge, documentation et fichiers source disponibles
'   sur le site <https://github.com/VincentChtln/NotaComp>
'
'   V1      Version initiale
'   V1.1    Amelioration de plusieurs fonctionnalit�s
'   V2      Refonte graphique et fonctionnelle
'   V2.4    Ajout de UserForm pour la modification des listes de classes
'   V2.5    Ajout de UserForm pour la modification des �valuations, modifications graphiques, am�lioration fonctionnelle du code
'
' *******************************************************************************
'                       GNU General Public License V3
'   Copyright (C)
'   Date: 2021
'   Auteur: Vincent Chatelain
'
'   This file is part of NotaComp.
'
'   NotaComp is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   NotaComp is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with NotaComp. If not, see <https://www.gnu.org/licenses/>.
'
'
'               GNU General Public License V3 - Traduction fran�aise
'
'   Ce fichier fait partie de NotaComp.
'
'   NotaComp est un logiciel libre ; vous pouvez le redistribuer ou le modifier
'   suivant les termes de la GNU General Public License telle que publi�e par la
'   Free Software Foundation, soit la version 3 de la Licence, soit (� votre gr�)
'   toute version ult�rieure.
'
'   NotaComp est distribu� dans l�espoir qu�il sera utile, mais SANS AUCUNE
'   GARANTIE : sans m�me la garantie implicite de COMMERCIALISABILIT�
'   ni d�AD�QUATION � UN OBJECTIF PARTICULIER. Consultez la GNU
'   General Public License pour plus de d�tails.
'
'   Vous devriez avoir re�u une copie de la GNU General Public License avec NotaComp;
'   si ce n�est pas le cas, consultez : <http://www.gnu.org/licenses/>.
'
' *******************************************************************************


' *******************************************************************************
'                               Module 1 - Accueil
'
'   Fonctions publiques
'
'   Proc�dures publiques
'
'   Fonctions priv�es
'
'   Proc�dures priv�es
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Constantes
' *******************************************************************************

' *** LOGICIEL ***
       Const strVersion         As String = "v2.5 - R�vision 2021"
       Const strLienGithub      As String = "https://github.com/VincentChtln/NotaComp"
       Const strLienSocleCommun As String = "https://www.education.gouv.fr/bo/15/Hebdo17/MENE1506516D.htm?cid_bo=87834"

' *** DONNEES PROTECTION ***
Global Const strPage1           As String = "Accueil"
Global Const strPage2           As String = "Listes"
Global Const strPassword        As String = vbNullString

' *** COULEURS ***
       Const byCouleurLogiciel  As Byte = 22
       Const byCouleurInfos     As Byte = 15
Global Const byCouleurCompet_1  As Byte = 50
Global Const byCouleurCompet_2  As Byte = 35
Global Const byCouleurClasse    As Byte = 44
Global Const byCouleurEval_1    As Byte = 45
Global Const byCouleurEval_2    As Byte = 40
Global Const byCouleurNote_1    As Byte = 42
Global Const byCouleurNote_2    As Byte = 34
Global Const byCouleurBilan     As Byte = 45

' *** LIMITES MIN ET MAX ***
       Const byNbClasses_Min    As Byte = 1
       Const byNbClasses_Max    As Byte = 20
       Const byNbEleves_Min     As Byte = 5
       Const byNbEleves_Max     As Byte = 50

' *** LIGNES ET COLONNES REFERENCE ***
       Const byLigTabLogiciel   As Byte = 5
       Const byColTabLogiciel   As Byte = 2
       Const byLigTabInfos      As Byte = byLigTabLogiciel + 5
       Const byColTabInfos      As Byte = 2
Global Const byLigTabClasses    As Byte = byLigTabInfos + 6
       Const byColTabClasses    As Byte = 2
Global Const byColTabCompet     As Byte = 2
Global Const byLigListePage2    As Byte = 1
Global Const byLigListePage3    As Byte = 6
Global Const byLigListePage4    As Byte = 3

' *** NOTES REFERENCE ***
Global Const dNoteA_Min         As Double = 3.5
Global Const dNoteB_Min         As Double = 2.5
Global Const dNoteC_Min         As Double = 1.5
Global Const dNoteD_Min         As Double = 0

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

Public Function isWorkbookProtected() As Boolean
    isWorkbookProtected = ThisWorkbook.ProtectWindows Or ThisWorkbook.ProtectStructure
End Function

Public Function isWorksheetProtected(ByVal ws As Worksheet) As Boolean
    isWorksheetProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
End Function

Public Function getNomPage3(ByVal byClasse As Byte) As String
    getNomPage3 = "Notes (" & getNomClasse(byClasse) & ")"
End Function

Public Function getNomPage4(ByVal byClasse As Byte) As String
    getNomPage4 = "Bilan (" & getNomClasse(byClasse) & ")"
End Function

Public Function getNomClasse(ByVal byClasse As Byte) As String
    getNomClasse = ThisWorkbook.Worksheets(strPage1).Cells(byLigTabClasses + byClasse + 2, byColTabClasses).Value
End Function

Public Function getIndiceClasse(strNomWs As String) As Byte
    Dim byClasse As Byte
    
    getIndiceClasse = 0
    
    For byClasse = 1 To getNombreClasses
        If InStr(strNomWs, getNomClasse(byClasse)) <> 0 Then getIndiceClasse = byClasse
    Next byClasse
End Function

Public Function getNombreClasses() As Byte
    getNombreClasses = CByte(ThisWorkbook.Worksheets(strPage1).Cells(byLigTabClasses, byColTabClasses + 1).Value)
End Function

Public Function getNombreEleves(ByVal byClasse As Byte) As Byte
    getNombreEleves = CByte(ThisWorkbook.Worksheets(strPage1).Cells(byLigTabClasses + byClasse + 2, byColTabClasses + 1).Value)
End Function

Public Function getTailleArray(ByRef arr As Variant) As Byte
    Dim byTaille        As Byte
    Dim byDimension     As Byte
    byTaille = 1
    byDimension = 1
    getTailleArray = 0
    
    On Error GoTo ErrorHandler
    Do While True
        byTaille = byTaille * (UBound(arr, byDimension) - LBound(arr, byDimension) + 1)
        If (byTaille > 1) Or Not (IsEmpty(arr(1))) Then getTailleArray = byTaille
        byDimension = byDimension + 1
    Loop
    Exit Function
    
ErrorHandler:
    If Err.Number = 13 Then ' Type Mismatch Error
        Err.Raise vbObjectError, "getTailleArray" _
            , "The argument passed to the getTailleArray function is not an array."
    End If
End Function

Public Function getTailleJaggedArray(ByRef arr As Variant) As Byte
    Dim iTailleExter As Byte
    Dim iElement As Byte
    
    iTailleExter = getTailleArray(arr)
    getTailleJaggedArray = 0
    
    For iElement = 1 To iTailleExter
        getTailleJaggedArray = getTailleJaggedArray + getTailleArray(arr(iElement))
    Next iElement
End Function

Public Function getArrayDomaines() As String()
    Dim arrDomaines(1 To 8, 1 To 2) As String
    
    ' *** NOM COMPLET ***
    arrDomaines(1, 1) = "Domaine 1: Les langages pour penser et communiquer - " & vbNewLine & _
                        "Composante 1: La langue fran�aise"
    arrDomaines(2, 1) = "Domaine 1: Les langages pour penser et communiquer - " & vbNewLine & _
                        "Composante 2: Les langues vivantes �trang�res ou r�gionales"
    arrDomaines(3, 1) = "Domaine 1: Les langages pour penser et communiquer - " & vbNewLine & _
                        "Composante 3: Les langages math�matiques, scientifiques et informatiques"
    arrDomaines(4, 1) = "Domaine 1: Les langages pour penser et communiquer - " & vbNewLine & _
                        "Composante 4: Les langages des arts et du corps"
    arrDomaines(5, 1) = "Domaine 2: Les m�thodes et outils pour apprendre"
    arrDomaines(6, 1) = "Domaine 3: La formation de la personne et du citoyen"
    arrDomaines(7, 1) = "Domaine 4: les syst�mes naturels et les syst�mes techniques"
    arrDomaines(8, 1) = "Domaine 5: Les repr�sentations du monde et de l'activit� humaine"
    
    ' *** ABREVIATION ***
    arrDomaines(1, 2) = "D1-1"
    arrDomaines(2, 2) = "D1-2"
    arrDomaines(3, 2) = "D1-3"
    arrDomaines(4, 2) = "D1-4"
    arrDomaines(5, 2) = "D2"
    arrDomaines(6, 2) = "D3"
    arrDomaines(7, 2) = "D4"
    arrDomaines(8, 2) = "D5"
    
    getArrayDomaines = arrDomaines
End Function

Public Function getArrayChoixCompetences() As Variant()
    ' *** DECLARATION VARIABLES ***
    Dim arrTamponSrc As Variant
    Dim arrTamponDest() As Variant
    Dim arrCompet() As Variant
    Dim arrChoixCompet(1 To 8) As Variant
    Dim byNbCompetParDomaine As Byte
    Dim byLigTabCompet As Byte
    Dim iCompetTampon As Byte
    Dim iCompetChoisie As Byte
    Dim iDrpValue As Byte
    Dim iDomaine As Byte
    Dim iCompet As Byte

    With ThisWorkbook.Worksheets(strPage1)
        ' *** AFFECTATION VARIABLES ***
        byLigTabCompet = byLigTabClasses + getNombreClasses + 7
        iDrpValue = .DropDowns("drpChoixCycle").Value
        iCompetTampon = 1
        arrCompet = getArrayCompetences(iDrpValue + 1)
        arrTamponSrc = .Range(.Cells(byLigTabCompet + 1, byColTabCompet + 2), _
                              .Cells(byLigTabCompet + getNombreCompetences(iDrpValue + 1), byColTabCompet + 3))
                              
        ' *** BOUCLE SUR TOUS LES DOMAINES ET COMPETENCES ***
        For iDomaine = 1 To 8
            byNbCompetParDomaine = getTailleArray(arrCompet(iDomaine))
            iCompetChoisie = 1
            ReDim arrTamponDest(1 To 1)
            For iCompet = 1 To byNbCompetParDomaine
            
                ' *** SI 'x' OU 'X' DANS LE TABLEAU, ALORS COPIE ABREVIATION COMPETENCE ***
                If arrTamponSrc(iCompetTampon, 1) = "X" Or arrTamponSrc(iCompetTampon, 1) = "x" Then
                    ReDim Preserve arrTamponDest(1 To iCompetChoisie)
                    arrTamponDest(iCompetChoisie) = arrTamponSrc(iCompetTampon, 2)
                    iCompetChoisie = iCompetChoisie + 1
                End If
                iCompetTampon = iCompetTampon + 1
            Next iCompet
            arrChoixCompet(iDomaine) = arrTamponDest
        Next iDomaine
    End With
    getArrayChoixCompetences = arrChoixCompet
End Function

' *******************************************************************************
'                               Proc�dures publiques
' *******************************************************************************

Public Sub protectWorkbook()
    ThisWorkbook.Protect Password:=strPassword, Structure:=True, Windows:=True
End Sub

Public Sub unprotectWorkbook()
    ThisWorkbook.Unprotect Password:=strPassword
End Sub

Public Sub protectWorksheet(ByRef ws As Worksheet)
    ws.Protect Password:=strPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterFaceOnly:=True
    ws.EnableSelection = xlUnlockedCells
End Sub

Public Sub protectAllWorksheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        protectWorksheet ws
    Next ws
End Sub

Public Sub unprotectWorksheet(ByRef ws As Worksheet)
    ws.Unprotect Password:=strPassword
End Sub

Public Sub enableUpdates()
    With Application
        .ScreenUpdating = True
        .StatusBar = "Pr�t"
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .PrintCommunication = True
    End With
End Sub

Public Sub disableUpdates()
    With Application
        .ScreenUpdating = False
        .StatusBar = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
        .PrintCommunication = False
    End With
End Sub

Public Sub freezePanes(ByRef wdw As Window, ByVal byLig As Byte, ByVal byCol As Byte)
    With wdw
        .SplitRow = byLig
        .SplitColumn = byCol
        .freezePanes = True
    End With
End Sub

Public Sub unloadAllUserForms()
    Dim uf As UserForm
    
    For Each uf In VBA.UserForms
        If TypeOf uf Is UserForm Then Unload uf
    Next uf
End Sub

Public Sub deleteAllButtons(ByRef ws As Worksheet)
    ws.Buttons.Delete
End Sub

Public Sub addWorksheet(ByVal sNom As String)
    If isWorkbookProtected Then Exit Sub
    
    With ThisWorkbook
        .Worksheets.Add After:=.Worksheets(.Worksheets.Count)
        .Worksheets(.Worksheets.Count).Name = sNom
    
        With .Worksheets(sNom)
            With .Cells
                .Borders.ColorIndex = 2
                .Locked = True
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            
            If InStr(sNom, "Notes") Or InStr(sNom, "Bilan") Then
                .Columns.ColumnWidth = 10
            Else
                .Columns.ColumnWidth = 20
            End If
        End With
    End With
End Sub

Public Sub creerTableau(ByVal strNomWs As String, ByVal rngCelOrigine As Range, _
                        ByVal iHaut As Byte, ByVal iLarg As Byte, ByVal iOrientation As Byte, _
                        ByRef arrAttribut() As String, ByVal byCouleur As Byte, Optional ByVal bLocked As Boolean = True)
    ' ***  DECLARATION VARIABLES ***
    Dim ws As Worksheet
    Dim bWsNomOK As Boolean
    Dim iAttribut As Byte
    
    ' *** VERIFICATIONS DONNEES ENTREES ***
    bWsNomOK = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = strNomWs Then
            bWsNomOK = True
            Exit For
        End If
    Next ws
    If Not (bWsNomOK) Or Not (rngCelOrigine.Count = 1) Or Not (iHaut >= 1) Or Not (iLarg >= 1) Then Exit Sub
    If Not (iOrientation = 1) And Not (iOrientation = 2) Then Exit Sub
    If (iOrientation = 1) And (Not (iHaut >= 2) Or Not (getTailleArray(arrAttribut) = iLarg)) Then Exit Sub
    If (iOrientation = 2) And (Not (iLarg >= 2) Or Not (getTailleArray(arrAttribut) = iHaut)) Then Exit Sub
    If Not ((byCouleur >= 1) And (byCouleur <= 56)) Then Exit Sub
    
    ' *** CREATION TABLEAU ***
    With ThisWorkbook.Worksheets(strNomWs)
        With .Range(rngCelOrigine.Address, .Cells(rngCelOrigine.Row + iHaut - 1, rngCelOrigine.Column + iLarg - 1))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = 1
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlHAlignCenter
        End With
        Select Case iOrientation
            ' *** TAB VERTICAL --> ATTRIBUTS EN HAUT ***
        Case 1
            .Range(rngCelOrigine.Address, .Cells(rngCelOrigine.Row, rngCelOrigine.Column + iLarg - 1)).Interior.ColorIndex = byCouleur
            For iAttribut = LBound(arrAttribut) To UBound(arrAttribut)
                .Cells(rngCelOrigine.Row, rngCelOrigine.Column + iAttribut - LBound(arrAttribut)).Value = arrAttribut(iAttribut)
            Next iAttribut
            .Range(.Cells(rngCelOrigine.Row + 1, rngCelOrigine.Column), _
                   .Cells(rngCelOrigine.Row + iHaut - 1, rngCelOrigine.Column + iLarg - 1)).Locked = bLocked
            
            ' *** TAB HORIZONTAL --> ATTRIBUTS SUR LE COTE ***
        Case 2
            .Range(rngCelOrigine.Address, .Cells(rngCelOrigine.Row + iHaut - 1, rngCelOrigine.Column)).Interior.ColorIndex = byCouleur
            For iAttribut = LBound(arrAttribut) To UBound(arrAttribut)
                .Cells(rngCelOrigine.Row + iAttribut - LBound(arrAttribut), rngCelOrigine.Column).Value = arrAttribut(iAttribut)
            Next iAttribut
            .Range(.Cells(rngCelOrigine.Row, rngCelOrigine.Column + 1), _
                   .Cells(rngCelOrigine.Row + iHaut - 1, rngCelOrigine.Column + iLarg - 1)).Locked = bLocked
        End Select
    End With
End Sub

' *******************************************************************************
'                               Fonctions priv�es
' *******************************************************************************
Private Function isInfosOK() As Boolean
    isInfosOK = False
    With ThisWorkbook.Worksheets(strPage1)
        If WorksheetFunction.CountBlank(.Range(.Cells(byLigTabInfos, byColTabInfos + 1), .Cells(byLigTabInfos + 3, byColTabInfos + 1))) = 0 Then isInfosOK = True
    End With
End Function

Private Function isNbClassesOK() As Boolean
    isNbClassesOK = False
    If getNombreClasses <> -1 Then isNbClassesOK = True
End Function

Private Function isNbEleveOK(ByVal varNbEleve As Variant) As Boolean
    isNbEleveOK = False
    If IsNumeric(varNbEleve) Then
        If varNbEleve > byNbEleves_Min And varNbEleve < byNbEleves_Max Then isNbEleveOK = True
    End If
End Function

Private Function isDonneesClassesOK() As Boolean
    ' *** DECLARATION VARIABLES ***
    Dim byNbClasses As Byte
    Dim byLigFinTableauClasses As Byte
    Dim rngTableauClasses As Range
    Dim celDonnee As Range
    
    isDonneesClassesOK = False
    byNbClasses = getNombreClasses
    If byNbClasses = -1 Then GoTo EOF
    
    With ThisWorkbook.Worksheets(strPage1)
        ' *** VERIFICATION NOMBRE CLASSES ***
        byLigFinTableauClasses = .Range(Mid(.UsedRange.Address, InStr(1, .UsedRange.Address, ":") + 1)).Row
        If byLigTabClasses + byNbClasses + 2 <> byLigFinTableauClasses Then GoTo EOF
        
        ' *** VERIFICATION VALEUR MANQUANTE ***
        Set rngTableauClasses = .Range(.Cells(byLigTabClasses + 3, byColTabClasses), .Cells(byLigFinTableauClasses, byColTabClasses + 1))
        If WorksheetFunction.CountBlank(rngTableauClasses) <> 0 Then GoTo EOF
        
        ' *** VERIFICATION NOMBRE ELEVES ***
        For Each celDonnee In rngTableauClasses
            If celDonnee.Column = byColTabClasses + 1 Then
                If Not (isNbEleveOK(celDonnee.Value)) Then GoTo EOF
            End If
        Next celDonnee
    End With
    
    isDonneesClassesOK = True
    
EOF:
End Function

Private Function getArrayCompetences(ByVal iCycle As Byte) As Variant()
    Dim arrCompetencesCycleI() As Variant
    
    ReDim arrCompetencesCycleI(1 To 8)
    Select Case iCycle
    
        ' *** COMPETENCES CYCLE 2 ***
    Case 2
        arrCompetencesCycleI(1) = Array("C1 - Comprendre et s'exprimer � l'oral", _
                                        "C2 - Lire et comprendre l'�crit", _
                                        "C3 - Ecrire", _
                                        "C4 - Utiliser � bon escient les r�gularit�s qui organisent la langue fran�aise " & _
                                        "(dans la limite de celles qui ont �t� �tudi�es)")
        arrCompetencesCycleI(2) = Array("C1 - Comprendre � l'oral (et � l'�crit)", _
                                        "C2 - S'exprimer � l'oral")
        arrCompetencesCycleI(3) = Array("C1 - Utiliser les nombres entiers", _
                                        "C2 - Reconna�tre des solides usuels et des figures g�om�triques", _
                                        "C3 - Se rep�rer et se d�placer")
        arrCompetencesCycleI(4) = Array("C1 - S'exprimer par des activit�s physiques, sportives ou artistiques, impliquant le corps", _
                                        "C2 - Partager et comprendre les langages artistiques")
        arrCompetencesCycleI(5) = Array("C1 - Organiser son travail personnel", _
                                        "C2 - Coop�rer avec des pairs", _
                                        "C3 - Rechercher et trairer l'information au moyen d'outils num�riques")
        arrCompetencesCycleI(6) = Array("C1 - S'exprimer (�motions, opinions, pr�f�rences) et respecter l'expression d'autrui", _
                                        "C2 - Prendre en compte les r�gles communes", _
                                        "C3 - Manifester son appartenance � un collectif")
        arrCompetencesCycleI(7) = Array("C1 - R�soudre des probl�mes �l�mentaires", _
                                        "C2 - Mener quelques �tapes d'une d�marche scientifique" & _
                                        "C3 - Mettre en pratique des comportements simples respectueux des autres, de l'environnement, de sa sant�")
        arrCompetencesCycleI(8) = Array("C1 - Se situer dans le temps et l'espace", _
                                        "C2 - Analyser et comprendre des organisations humaines et les repr�sentations du monde", _
                                        "C3 - Imaginer, �laborer, produire")
        
        ' *** COMPETENCES CYCLE 3 ***
    Case 3
        arrCompetencesCycleI(1) = Array("C1 - S'exprimer � l'oral", _
                                        "C2 - Comprendre des �nonc�s oraux", _
                                        "C3 - Lire et comprendre l'�crit", _
                                        "C4 - Ecrire", _
                                        "C5 - Exploiter les ressources de la langues / R�fl�chir sur le syst�me linguistique")
        arrCompetencesCycleI(2) = Array("C1 - Lire et comprendre l'�crit", _
                                        "C2 - Ecrire et r�agir � l'�crit", _
                                        "C3 - Ecouter et comprendre", _
                                        "C4 - S'exprimer � l'oral en continu et en interaction")
        arrCompetencesCycleI(3) = Array("C1 - Utiliser les nombres entiers, les nombres d�cimaux et les fractions simples", _
                                        "C2 - Reconna�tre des solides usuels et des figures g�om�triques", _
                                        "C3 - Se rep�rer et se d�placer")
        arrCompetencesCycleI(4) = Array("C1 - S'exprimer par des activit�s physiques sportives ou artistiques", _
                                        "C2 - Pratiquer des arts en mobilisant divers langages artistiques et leurs ressources expressives / " & _
                                        "Prendre du recul sur la pratique artistique individuelle et collective")
        arrCompetencesCycleI(5) = Array("C1 - Se constituter des outils de travail personnel et mettre en place des strat�gies pour comprendre et apprendre", _
                                        "C2 - Coop�rer et r�aliser des projets", _
                                        "C3 - Rechercher et trairer l'information et s'initier aux langages des m�dias", _
                                        "C4 - Mobiliser des outils num�riques pour apprendre, �changer et communiquer")
        arrCompetencesCycleI(6) = Array("C1 - Ma�triser l'expression de sa sensibilit� et de ses opinions, respecter celles des autres", _
                                        "C2 - Comprendre et conna�tre la r�gle et le droit", _
                                        "C3 - Exercer son esprit critique, faire preuve de r�flexion et de discernement")
        arrCompetencesCycleI(7) = Array("C1 - Mener une d�marche scientifique ou technologique, r�soudre des probl�mes simples", _
                                        "C2 - Mettre en pratique des comportements simples respectueux des autres, de l'environnement, de sa sant�")
        arrCompetencesCycleI(8) = Array("C1 - Se situer dans le temps et l'espace", _
                                        "C2 - Analyser et comprendre des organisations humaines et les repr�sentations du monde", _
                                        "C3 - Raisonner, imaginer, �laborer, produire")

        ' *** COMPETENCES CYCLE 4 ***
    Case 4
        arrCompetencesCycleI(1) = Array("C1 - S'exprimer � l'oral", _
                                        "C2 - Comprendre des �nonc�s oraux", _
                                        "C3 - Lire et comprendre l'�crit", _
                                        "C4 - Ecrire", _
                                        "C5 - Exploiter les ressources de la langues / R�fl�chir sur le syst�me linguistique")
        arrCompetencesCycleI(2) = Array("C1 - Lire et comprendre l'�crit", _
                                        "C2 - Ecrire et r�agir � l'�crit", _
                                        "C3 - Ecouter et comprendre", _
                                        "C4 - S'exprimer � l'oral en continu et en interaction")
        arrCompetencesCycleI(3) = Array("C1 - Utiliser les nombres", _
                                        "C2 - Utiliser un calcul litt�ral", _
                                        "C3 - Exprimer une grandeur mesur�e ou calcul�e dans une unit� adapt�e", _
                                        "C4 - Passer d'un langage � un autre", _
                                        "C5 - Utiliser le langage des probabilit�s", _
                                        "C6 - Utiliser et produire des repr�sentations d'objets", _
                                        "C7 - Utiliser l'algorithmique et la programmation pour cr�er des applications simples")
        arrCompetencesCycleI(4) = Array("C1 - Pratiquer des activit�s physiques sportives et     artistiques", _
                                        "C2 - Pratiquer des arts en mobilisant divers langages artistiques et leurs ressources expressives / " & _
                                        "Prendre du recul sur la pratique artistique individuelle et collective")
        arrCompetencesCycleI(5) = Array("C1 - Organiser son travail personnel", _
                                        "C2 - Coop�rer et r�aliser des projets", _
                                        "C3 - Rechercher et trairer l'information et s'initier aux langages des m�dias", _
                                        "C4 - Mobiliser des outils num�riques pour apprendre, �changer et communiquer")
        arrCompetencesCycleI(6) = Array("C1 - Ma�triser l'expression de sa sensibilit� et de ses opinions, respecter celles des autres", _
                                        "C2 - Comprendre et conna�tre la r�gle et le droit", _
                                        "C3 - Exercer son esprit critique, faire preuve de r�flexion et de discernement", _
                                        "C4 - Faire preuve de responsabilit�, respecter les r�gles de vie collective, s'engager et prendre des initiatives")
        arrCompetencesCycleI(7) = Array("C1 - Mener une d�marche scientifique", _
                                        "C2 - Concevoir des objets et syst�mes techniques", _
                                        "C3 - Identifier des r�gles et des principes de responsabilit� individuelle et collective " & _
                                        "dans le domaine de la sant�, de la s�curit�, de l'environnement")
        arrCompetencesCycleI(8) = Array("C1 - Se situer dans le temps et l'espace", _
                                        "C2 - Analyser et comprendre des organisations humaines et les repr�sentations du monde", _
                                        "C3 - Raisonner, imaginer, �laborer, produire")
        
        ' *** VALEUR iCycle NON VALIDE
    Case Else
        ReDim arrCompetencesCycleI(1 To 1)
        arrCompetencesCycleI = Array("")
    End Select
    
    getArrayCompetences = arrCompetencesCycleI

End Function

Private Function getNombreCompetences(ByVal iCycle As Byte) As Byte
    getNombreCompetences = getTailleJaggedArray(getArrayCompetences(iCycle))
End Function

Private Function isCompetOK() As Boolean
    isCompetOK = False
    If True Then isCompetOK = True
End Function

' *******************************************************************************
'                               Proc�dures priv�es
' *******************************************************************************

Private Sub initNotaComp()
    ' *** DECLARATION VARIABLES ***
    Dim rngBtnDemarrerNotacomp As Range
    Dim btnDemarrerNotacomp As Variant
    Dim shp As Shape
    
    ' *** REFRESH ECRAN OFF ***
    disableUpdates

    ' *** VERIFICATION PROTECTION WORKBOOK ***
    If isWorkbookProtected Then
        MsgBox "Ce classeur est prot�g�. Enlevez la protection avant de continuer."
        GoTo EOP
    End If
        
    With ThisWorkbook
        ' *** VERIFICATION WORKSHEET UNIQUE ***
        If .Worksheets.Count <> 1 Then
            MsgBox "Ce classeur contient plusieurs feuilles. Supprimez toutes les feuilles sauf une avant de continuer."
            GoTo EOP
        End If
        
        ' *** VERIFICATION PROTECTION FEUILLE ***
        If isWorksheetProtected(.Worksheets(1)) Then
            MsgBox "Cette feuille est prot�g�e. Enlevez la protection avant de continuer."
            GoTo EOP
        End If
            
        With .Worksheets(1)
            ' *** VERIFICATION FEUILLE VIDE ***
            If Not (.UsedRange.Address = "$A$1" And .Range("A1") = vbNullString) Then
                Select Case MsgBox("La feuille n'est pas vide, cliquez sur 'OK' pour supprimer son contenu" & vbNewLine & _
                                   "ou sur 'Annuler' pour revenir en arri�re", vbOKCancel)
                Case vbOK
                    .Cells.Clear
                    .EnableSelection = xlUnlockedCells
                    For Each shp In .Shapes
                        shp.Delete
                    Next shp
                Case vbCancel
                    GoTo EOP
                End Select
            End If
            
            ' *** MISE EN PAGE ***
            .Name = strPage1
            .Rows().RowHeight = 20
            .Rows(2).RowHeight = 50
            .Columns().ColumnWidth = 30
            .Columns(1).ColumnWidth = 5
            .Columns(2).ColumnWidth = 50
            .Columns(3).ColumnWidth = 50
            .Columns(4).ColumnWidth = 20
            With .Cells
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlHAlignCenter
                .Borders.ColorIndex = 2
                .Locked = True
            End With
            With .Range("B2")
                .HorizontalAlignment = xlHAlignLeft
                .Font.Size = 40
                .Value = "NotaComp"
            End With
            
            ' *** AJOUT BOUTON 'DEMARRER NOTACOMP' ***
            Set rngBtnDemarrerNotacomp = .Range(.Cells(byLigTabLogiciel, byColTabLogiciel + 3).Address)
            Set btnDemarrerNotacomp = .Buttons.Add(rngBtnDemarrerNotacomp.Left, rngBtnDemarrerNotacomp.Top, _
                                                   rngBtnDemarrerNotacomp.Width, rngBtnDemarrerNotacomp.Height)
            With btnDemarrerNotacomp
                .Caption = "D�marrer NotaComp"
                .OnAction = "btnDemarrerConfiguration_Click"
            End With

            ' *** AJOUT TABLEAU INFOS LOGICIEL ***
            creerTableauLogiciel
        End With
        ' *** PROTECTION + REFRESH ECRAN ON ***
        protectWorksheet .Worksheets(strPage1)
        protectWorkbook
    End With
    
EOP:
    enableUpdates
End Sub

Private Sub creerTableauLogiciel()
    Dim arrAttributLogiciel(1 To 3) As String
    
    arrAttributLogiciel(1) = "Version de l'outil"
    arrAttributLogiciel(2) = "Classeur vierge, documents et code source"
    arrAttributLogiciel(3) = "Textes officiels - Socle commun"
    
    With ThisWorkbook.Worksheets(strPage1)
        creerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabLogiciel, byColTabLogiciel), _
        iHaut:=3, iLarg:=2, iOrientation:=2, _
        arrAttribut:=arrAttributLogiciel, byCouleur:=byCouleurLogiciel
        .Cells(byLigTabLogiciel, byColTabLogiciel + 1).Value = strVersion
        .Hyperlinks.Add Anchor:=.Cells(byLigTabLogiciel + 1, byColTabLogiciel + 1), _
        Address:=strLienGithub, TextToDisplay:="D�p�t Github"
        .Hyperlinks.Add Anchor:=.Cells(byLigTabLogiciel + 2, byColTabLogiciel + 1), _
        Address:=strLienSocleCommun, TextToDisplay:="Bulletin Officiel - D�cret n� 2015-372 du 31/03/2015"
        .Range(.Cells(byLigTabLogiciel + 1, byColTabLogiciel + 1), .Cells(byLigTabLogiciel + 2, byColTabLogiciel + 1)).Locked = False
    End With
End Sub

Private Sub btnDemarrerConfiguration_Click()
    ' *** REFRESH ECRAN OFF ***
    disableUpdates
    
    ' *** MESSAGE INFORMATION ***
    MsgBox "Bienvenue dans la configuration de NotaComp. Prennez le temps n�cessaire pour r�aliser correctement " & _
           "cette d�marche, car il ne sera pas possible de revenir dessus plus tard. " & _
           "Le processus est d�coup� en plusieurs �tapes, et des messages comme celui-ci " & _
           "s'afficheront pour vous donner des indications. Merci de les lire attentivement." & vbNewLine & vbNewLine & _
           "Au fur et � mesure, des informations vous seront demand�es concernant:" & vbNewLine & _
           "    - Les classes auxquelles vous enseignez, avec la liste des �l�ves" & vbNewLine & _
           "    - Les comp�tences �valu�es au cours de l'ann�e (selon la d�nomination officielle)" & vbNewLine & vbNewLine & _
           "Pensez � pr�parer ces �l�ments au pr�alable pour faciliter la configuration." & vbNewLine & vbNewLine & _
           "C'est parti !"
    ' *** SUPPRESSION BOUTONS ***
    deleteAllButtons ThisWorkbook.Worksheets(strPage1)
    
    ' *** APPEL PROCEDURE ***
    creerTableauInformations
    creerTableauNombreClasses

    ' *** REFRESH ECRAN ON ***
    enableUpdates
        
    ' *** MESSAGE INFORMATION ***
    MsgBox "Entrez tout d'abord vos informations dans le tableau d'informations (tableau gris)." & vbNewLine & _
           "Puis entrez le nombre de classes dans la case correspondante (tableau jaune). " & _
           "Cliquez ensuite sur le bouton 'Valider le nombre de classes' pour passer � l'�tape suivante." & vbNewLine & vbNewLine & _
           "ATTENTION: ce classeur peut comporter des classes de diff�rents niveaux, mais appartenant toutes au m�me cycle (2, 3 ou 4). " & _
           "Par exemple, si vous enseignez en coll�ge � des classes de Sixi�me et de Quatri�me, il vous faudra faire deux classeurs s�par�s."

End Sub

Private Sub creerTableauInformations()
    Dim arrAttributInfos(1 To 4) As String
    
    arrAttributInfos(1) = "Etablissement"
    arrAttributInfos(2) = "Mati�re"
    arrAttributInfos(3) = "Professeur"
    arrAttributInfos(4) = "Ann�e scolaire"
    
    With ThisWorkbook.Worksheets(strPage1)
        creerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabInfos, byColTabInfos), _
        iHaut:=4, iLarg:=2, iOrientation:=2, _
        arrAttribut:=arrAttributInfos, byCouleur:=byCouleurInfos, bLocked:=False
    End With
End Sub

Private Sub creerTableauNombreClasses()
    ' *** DECLARATION VARIABLES ***
    Dim arrAttributClasses(1 To 1) As String
    Dim rngBtnValiderNbClasses As Range
    Dim btnValiderNbClasses As Variant
    
    arrAttributClasses(1) = "Nombre de classes"
    
    With ThisWorkbook.Worksheets(strPage1)
        ' *** CREATION TABLEAU ***
        creerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabClasses, byColTabClasses), _
        iHaut:=1, iLarg:=2, iOrientation:=2, _
        arrAttribut:=arrAttributClasses, byCouleur:=byCouleurClasse, bLocked:=False

        ' *** AJOUT BOUTON 'VALIDER NOMBRE CLASSES' ***
        Set rngBtnValiderNbClasses = .Range(.Cells(byLigTabClasses, byColTabClasses + 3).Address)
        With .Buttons.Add(rngBtnValiderNbClasses.Left, rngBtnValiderNbClasses.Top, _
                          rngBtnValiderNbClasses.Width, rngBtnValiderNbClasses.Height)
            .Caption = "Valider le nombre de classes"
            .OnAction = "btnValiderNombreClasses_Click"
            .Name = "btnValiderNombreClasses"
        End With
        Set rngBtnValiderNbClasses = Nothing
    End With
End Sub

Private Sub btnValiderNombreClasses_Click()
    ' *** DECLARATION VARIABLES ***
    Dim byNbClasses As Byte
    
    ' *** REFRESH ECRAN OFF ***
    disableUpdates
    
    ' *** VERIFICATION VALEUR ***
    If Not (isNbClassesOK) Then
        MsgBox "ATTENTION: le nombre de classes n'est pas valide  (nombre min = " & byNbClasses_Min & ", nombre max = " & byNbClasses_Max & ")."
        GoTo EOP
    Else
        byNbClasses = getNombreClasses
    End If
    
    ' *** MODIFICATION BOUTON ***
    With ThisWorkbook.Worksheets(strPage1).Buttons("btnValiderNombreClasses")
        .LockedText = False
        .Caption = "Modifier le nombre de classes"
        .OnAction = "btnModifierNombreClasses_Click"
        .Name = "btnModifierNombreClasses"
    End With
    
    ' *** APPEL PROCEDURE ***
    creerTableauClasses byNbClasses
    
    ' *** REFRESH ECRAN ON ***
    enableUpdates

    ' *** MESSAGE INFORMATION ***
    MsgBox "Entrez maintenant le nom de chaque classe ainsi que le nombre d'�l�ves qui s'y trouve. " & _
           "Cela permettra au tableur de g�n�rer les listes d'�l�ves � compl�ter." & vbNewLine & vbNewLine & _
           "Si besoin, vous pouvez changer le nombre de classes � l'aide du bouton 'Modifier'. " & _
           "Une fois toutes les cases compl�t�es, cliquez sur le bouton 'Valider' pour bloquer les donn�es et passer � l'�tape suivante." & vbNewLine & vbNewLine & _
           "INDICATION: Les noms de classes sont limit�s � 7 caract�res. Utilisez donc des noms courts, " & _
           "par exemple '5�me 2', '5�me2', '5e 2' ou encore '5e2' pour d�signer la classe de Cinqui�me 2."
    
    Exit Sub

EOP:
    ' *** REFRESH ECRAN ON ***
    enableUpdates
End Sub

Private Sub btnModifierNombreClasses_Click()
    ' *** DECLARATION VARIABLES ***
    Dim byNbClasses As Byte
    
    ' *** REFRESH ECRAN OFF ***
    disableUpdates
    
    ' *** VERIFICATION VALEUR ***
    If Not (isNbClassesOK) Then
        MsgBox "ATTENTION: le nombre de classes n'est pas valide  (nombre min = " & byNbClasses_Min & ", nombre max = " & byNbClasses_Max & ")."
        GoTo EOP
    Else
        byNbClasses = getNombreClasses
    End If
    
    ' *** SUPPRESSION CELLULES ***
    With ThisWorkbook.Worksheets(strPage1)
        .Range(.Cells(byLigTabClasses + 1, byColTabClasses), .Cells(byLigTabClasses + byNbClasses_Max + 10, byColTabClasses + 1)).Delete Shift:=xlUp
        .Buttons("btnValiderClasses").Delete
    End With

    ' *** APPEL PROCEDURE ***
    creerTableauClasses byNbClasses

EOP:
    ' *** REFRESH ECRAN ON ***
    enableUpdates
End Sub

Private Sub creerTableauClasses(ByVal byNbClasses As Byte)
    ' *** DECLARATION VARIABLES ***
    Dim arrAttributClasses(1 To 2) As String
    Dim rngBtnValiderClasses As Range
    Dim btnValiderClasses As Variant
    
    ' *** AFFECTATION VARIABLES ***
    arrAttributClasses(1) = "Nom de la classe"
    arrAttributClasses(2) = "Nombre d'�l�ves"
    
    With ThisWorkbook.Worksheets(strPage1)
        ' *** CREATION TABLEAU ***
        creerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabClasses + 2, byColTabClasses), _
        iHaut:=byNbClasses + 1, iLarg:=2, iOrientation:=1, _
        arrAttribut:=arrAttributClasses, byCouleur:=byCouleurClasse, bLocked:=False
        
        ' *** AJOUT BOUTON 'VALIDER CLASSES' ***
        Set rngBtnValiderClasses = .Range(.Cells(byLigTabClasses + 2, byColTabClasses + 3).Address)
        With .Buttons.Add(rngBtnValiderClasses.Left, rngBtnValiderClasses.Top, _
                          rngBtnValiderClasses.Width, rngBtnValiderClasses.Height)
            .Caption = "Valider les classes"
            .OnAction = "btnValiderClasses_Click"
            .Name = "btnValiderClasses"
        End With
        Set rngBtnValiderClasses = Nothing
    End With
End Sub

Private Sub btnValiderClasses_Click()
    ' ***REFRESH ECRAN OFF ***
    disableUpdates
    
    ' *** CONFIRMATION UTILISATEUR
    If Not (MsgBox("Confirmez-vous le nom de classe et le nombre d'�l�ves indiqu�s? " & _
                   "Il ne sera pas possible de les modifier par la suite.", vbYesNo) = vbYes) Then GoTo EOP
    
    ' *** VERIFICATION VALEUR ***
    If Not (isDonneesClassesOK) Then
        MsgBox "ATTENTION: les donn�es entr�es pour les classes ne sont pas valides, cela emp�che de passer � la prochaine �tape." & vbNewLine & _
               "Cela peut provenir de trois �l�ments:" & vbNewLine & _
               "   - Nombre de classes qui ne correspond pas � la taille du tableau," & vbNewLine & _
               "   - Donn�es manquantes dans le tableau," & vbNewLine & _
               "   - Nombre d'�l�ves incorrect (nombre min = " & byNbEleves_Min & ", nombre max = " & byNbEleves_Max & ")." & vbNewLine & _
               "V�rifiez ces trois propri�t�s et corrigez-les pour continuer."
        GoTo EOP
    End If
    
    ' *** SUPPRESSION BOUTONS ***
    With ThisWorkbook.Worksheets(strPage1)
        .Range(.Cells(byLigTabClasses, byColTabClasses), .Cells(byLigTabClasses + getNombreClasses + 2, byColTabClasses + 1)).Locked = True
        ThisWorkbook.Worksheets(strPage1).Buttons.Delete
    End With

    ' *** APPEL PROCEDURE ***
    creerDropdownCycle
    
    ' ***REFRESH ECRAN ON ***
    enableUpdates

    ' *** MESSAGE INFORMATION ***
    MsgBox "Choisissez maintenant le cycle d'�tude (2, 3 ou 4). Un tableau de comp�tences s'affichera alors en-dessous et changera en fonction du cycle choisi. " & vbNewLine & vbNewLine & _
           "S�lectionnez ensuite les comp�tences que vous �valuerez au cours de l'ann�e (en fonction de votre mati�re, et pour l'ensemble des classes que vous avez indiqu�). " & _
           "Pour cela, �crivez la lettre 'X' dans la case correspondant � la comp�tence not�e." & vbNewLine & _
           "Une fois toutes vos comp�tences s�lectionn�es, cliquez sur 'Valider les comp�tences' pour passer � l'�tape suivante." & vbNewLine & vbNewLine & _
           "INDICATION: Afin d'am�liorer l'affichage, chaque comp�tence sera indiqu�e par une abr�viation � la place de son nom complet. " & vbNewLine & _
           "En s�lectionnant une comp�tence, une abr�viation par d�faut vous sera propos�e. Cependant, si vous disposez d'une abr�viation qui vous est propre, " & _
           "vous �tes libre de l'�crire dans la case correspondante. Elle sera alors utilis�e � la place de l'abr�viation par d�faut" & vbNewLine & _
           "Comme pr�c�demment, utilisez des abr�viations courtes car la taille est limit�e � 7 caract�res."
    Exit Sub
    
EOP:
    ' ***REFRESH ECRAN ON ***
    enableUpdates
End Sub

Private Sub creerDropdownCycle()
    ' *** DECLARATION VARIABLES ***
    Dim byLigChoixCycle As Byte
    Dim rngDrpChoixCycle As Range
    Dim drpChoixCycle As Variant
    Dim rngBtnValiderCompetences As Range
    Dim btnValiderCompetences As Variant

    ' *** AFFECTATION VARIABLES ***
    byLigChoixCycle = byLigTabClasses + getNombreClasses + 5
    
    With ThisWorkbook.Worksheets(strPage1)
        ' *** MISE EN FORME ***
        With .Range(.Cells(byLigChoixCycle, byColTabCompet), _
                    .Cells(byLigChoixCycle, byColTabCompet + 1)).Borders
            .ColorIndex = 1
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Cells(byLigChoixCycle, byColTabCompet)
            .Interior.ColorIndex = byCouleurCompet_1
            .Value = "Choix du cycle"
        End With
        
        ' *** AJOUT COMBOBOX CHOIX CYCLE ***
        Set rngDrpChoixCycle = .Cells(byLigChoixCycle, byColTabCompet + 1)
        Set drpChoixCycle = .DropDowns.Add(rngDrpChoixCycle.Left, rngDrpChoixCycle.Top, _
                                           rngDrpChoixCycle.Width, rngDrpChoixCycle.Height)
        With drpChoixCycle
            .DropDownLines = 3
            .AddItem "Cycle 2", 1
            .AddItem "Cycle 3", 2
            .AddItem "Cycle 4", 3
            .Name = "drpChoixCycle"
            .OnAction = "drpChoixCycle_Change"
        End With
    
        ' *** AJOUT BOUTON 'VALIDER CLASSES' ***
        Set rngBtnValiderCompetences = .Cells(byLigChoixCycle, byColTabCompet + 3)
        Set btnValiderCompetences = .Buttons.Add(rngBtnValiderCompetences.Left, rngBtnValiderCompetences.Top, _
                                                 rngBtnValiderCompetences.Width, rngBtnValiderCompetences.Height)
        With btnValiderCompetences
            .Caption = "Valider les comp�tences"
            .OnAction = "btnValiderCompetences_Click"
            .Name = "btnValiderCompetences"
        End With
    End With
End Sub

Public Sub drpChoixCycle_Change()
    ' *** DECLARATION VARIABLES ***
    Dim iDrpValue As Byte
    Dim byLigTabCompet As Byte
    
    ' *** AFFECTATION VARIABLES ***
    iDrpValue = ThisWorkbook.Worksheets(strPage1).DropDowns("drpChoixCycle").Value
    byLigTabCompet = byLigTabClasses + getNombreClasses + 7
    
    ' *** REFRESH ECRAN OFF ***
    disableUpdates
    
    ' *** OPERATION ***
    If iDrpValue = 1 Or iDrpValue = 2 Or iDrpValue = 3 Then
        With ThisWorkbook.Worksheets(strPage1)
            .Range(.Cells(byLigTabCompet, byColTabCompet), .Cells(byLigTabCompet + 100, byColTabCompet)).EntireRow.Delete
        End With
        creerTableauChoixCompetences iDrpValue
    End If
    
    ' *** REFRESH ECRAN ON ***
    enableUpdates
End Sub

Private Sub creerTableauChoixCompetences(ByVal iDrpValue As Byte)
    ' *** DECLARATION VARIABLES ***
    Dim byLigTabCompet As Byte
    Dim byLigDomaine As Byte
    Dim byLigCompetence As Byte
    Dim arrAttributCompet(1 To 4) As String
    Dim arrDomaines As Variant
    Dim arrCompetences As Variant
    Dim iDomaine As Byte
    Dim iCompetence As Byte
    Dim byNbCompetences As Byte
    Dim rowCompetence As Range
    
    ' *** VALEUR ATTRIBUTS ***
    arrAttributCompet(1) = "Domaines / Composantes"
    arrAttributCompet(2) = "Comp�tences"
    arrAttributCompet(3) = "S�lection"
    arrAttributCompet(4) = "Abr�viation"

    ' *** AFFECTATION VARIABLES ***
    arrDomaines = getArrayDomaines()
    arrCompetences = getArrayCompetences(iDrpValue + 1)
    byLigTabCompet = byLigTabClasses + getNombreClasses + 7
    byLigCompetence = byLigTabCompet + 1
    byNbCompetences = getNombreCompetences(iDrpValue + 1)

    With ThisWorkbook.Worksheets(strPage1)
        ' *** AJOUT DOMAINES & COMPETENCES ***
        For iDomaine = 1 To 8
            .Cells(byLigCompetence, byColTabCompet).Value = arrDomaines(iDomaine, 1)
            byLigDomaine = byLigCompetence
            For iCompetence = LBound(arrCompetences(iDomaine)) To UBound(arrCompetences(iDomaine))
                .Cells(byLigCompetence, byColTabCompet + 1).Value = arrCompetences(iDomaine)(iCompetence)
                byLigCompetence = byLigCompetence + 1
            Next iCompetence
            .Range(.Cells(byLigDomaine, byColTabCompet), .Cells(byLigCompetence - 1, byColTabCompet)).MergeCells = True
        Next iDomaine
        
        ' *** MISE EN FORME ***
        creerTableau strNomWs:=strPage1, rngCelOrigine:=.Cells(byLigTabCompet, byColTabCompet), _
                     iHaut:=byNbCompetences + 1, iLarg:=4, iOrientation:=1, _
                     arrAttribut:=arrAttributCompet, byCouleur:=byCouleurCompet_1
                     
        .Range(.Cells(byLigTabCompet + 1, byColTabCompet + 2), _
               .Cells(byLigTabCompet + byNbCompetences, byColTabCompet + 3)).Locked = False
        With .Range(.Cells(byLigTabCompet + 1, byColTabCompet), _
                    .Cells(byLigTabCompet + byNbCompetences, byColTabCompet + 1))
            .HorizontalAlignment = xlHAlignLeft
            .WrapText = True
            .Rows.AutoFit
            For Each rowCompetence In .Rows
                rowCompetence.RowHeight = 20 / 14.4 * rowCompetence.RowHeight
            Next rowCompetence
        End With
    End With
End Sub

Private Sub btnValiderCompetences_Click()
    Dim byNbCompetences As Byte
    Dim byLigTabCompet As Byte

    If Not (isInfosOK) Then
        MsgBox "Il manque des donn�es dans le tableau d'informations (tableau gris). " & _
               "Merci de compl�ter tous les cases avant de passer � l'�tape suivante."
        Exit Sub
    End If
    
    If Not (MsgBox("Confirmez-vous la s�lection des comp�tences � �valuer? " & _
                   "Il ne sera pas possible de la modifier par la suite.", vbYesNo) = vbYes) Then Exit Sub

    If Not (isCompetOK) Then
        MsgBox "ATTENTION: votre choix de comp�tences n'est pas valide. Cela peut provenir de deux �l�ments: " & vbNewLine & _
               "    - Moins de deux comp�tences s�lectionn�es pour �valuation" & vbNewLine & _
               "    - Abr�viations utilis�es trop longues (limite max = 7 caract�res)" & vbNewLine & _
               "V�rifiez ces deux propri�t�s et corrigez-les pour continuer."
        Exit Sub
    End If
    
    With ThisWorkbook.Worksheets(strPage1)
        byLigTabCompet = byLigTabClasses + getNombreClasses + 7
        byNbCompetences = getNombreCompetences(.Shapes("drpChoixCycle").ControlFormat.Value + 1)
        
        disableUpdates
        .Cells.Font.Bold = False
        .Range(.Cells(byLigTabLogiciel + 1, byColTabLogiciel + 1).Address).Activate
        .Range(.Cells(byLigTabInfos, byColTabInfos + 1), .Cells(byLigTabInfos + 3, byColTabInfos + 1)).Locked = True
        .Range(.Cells(byLigTabCompet + 1, byColTabCompet + 2), .Cells(byLigTabCompet + byNbCompetences, byColTabCompet + 3)).Locked = True
        .Shapes("drpChoixCycle").ControlFormat.Enabled = False
        deleteAllButtons ThisWorkbook.Worksheets(strPage1)
        unprotectWorkbook
        addWorksheet (strPage2)
        creerTableauListeClasses
        protectAllWorksheets
        protectWorkbook
        enableUpdates
    End With
    
    MsgBox "Vous arrivez maintenant sur la page de gestion des listes de classes. " & _
           "Vous pouvez ici entrer le noms et pr�nom des �l�ves de chaque classe, dans les colonnes correspondantes. " & vbNewLine & _
           "Une fois termin�, cliquez sur le bouton 'Valider les listes' pour passer � l'�tape suivante." & vbNewLine & vbNewLine & _
           "INDICATION: Afin de g�rer d'�ventuelles modifications de classe en cours d'ann�e (arriv�e ou d�part d'un �l�ve), " & _
           "il est possible de d'ajouter, de transf�rer ou de supprimer un �l�ve en cliquant sur le bouton 'Modifier les listes' " & _
           "(qui s'affichera � la place du bouton 'Valider les listes')."
End Sub


