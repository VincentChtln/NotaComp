VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Transférer un élève "
   ClientHeight    =   6096
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3780
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ##################################
' UF Transfert d'élève
' ##################################

Option Explicit

' ##################################
' PROCÉDURES
' ##################################
' UserForm_Initialize()
' listboxSelectionClasseSource_Change()
' btnTransfererEleve_Click()
' ##################################

' Initialisation de l'UF
' Peuplement de listboxSelectionClasseSource
Private Sub UserForm_Initialize()
    ' *** DECLARATION VARIABLES ***
    Dim intNbClasse As Integer
    Dim intIndiceClasse As Integer
    Dim strNomClasse As String
    
    ' *** AFFECTATION VARIABLES ***
    intNbClasse = getNombreClasses()
    
    ' *** AJOUT CLASSES DANS LISTE ***
    For intIndiceClasse = 1 To intNbClasse
        strNomClasse = getNomClasse(intIndiceClasse)
        listboxSelectionClasseSource.AddItem strNomClasse
    Next intIndiceClasse
    
    ' *** INITIALISATION INDEX ***
    listboxSelectionClasseSource.ListIndex = 0
End Sub

' Modification de la liste Eleve en fonction de la classe sélectionnée
Private Sub listboxSelectionClasseSource_Change()
    ' *** DECLARATION VARIABLES ***
    Dim intIndiceClasseSource As Integer
    Dim intColonneClasseSource As Integer
    Dim intNbEleves As Integer
    Dim intIndiceEleve As Integer
    Dim strNomCompletEleve As String
    Dim intNbClasses As Integer
    Dim intIndiceClasseDest As Integer
    Dim strNomClasseDest As String

    ' *** AFFECTATION VARIABLES ***
    listboxSelectionEleve.Clear
    listboxSelectionClasseDest.Clear
    intNbEleves = getNombreEleves(intIndiceClasseSource)
    intNbClasses = getNombreClasses()
    intIndiceClasseSource = listboxSelectionClasseSource.ListIndex + 1
    intColonneClasseSource = 2 * intIndiceClasseSource - 1

    ' *** AJOUT ELEVES DANS LISTE ***
    For intIndiceEleve = 1 To intNbEleves
        strNomCompletEleve = Worksheets(strPage2).Cells(intLigListePage2 + intIndiceEleve, intColonneClasseSource).Value
        listboxSelectionEleve.AddItem strNomCompletEleve
    Next intIndiceEleve
    
    ' *** AJOUT CLASSES DANS LISTE TRANSFERT ***
    For intIndiceClasseDest = 1 To intNbClasses
        If intIndiceClasseDest <> intIndiceClasseSource Then
            strNomClasseDest = getNomClasse(intIndiceClasseDest)
            listboxSelectionClasseDest.AddItem strNomClasseDest
        End If
    Next intIndiceClasseDest

    ' *** INITIALISATION INDEX ***
    listboxSelectionEleve.ListIndex = 0
    listboxSelectionClasseDest.ListIndex = 0
End Sub

' Demande de confirmation puis appel de la procédure transfererEleve (Module 2)
Private Sub btnTransfererEleve_Click()
    ' *** DECLARATION VARIABLES ***
    Dim intClasseSource As Integer
    Dim strNomClasseSource As String
    Dim intClasseDest As Integer
    Dim strNomClasseDest As String
    Dim strNomCompletEleve As String
    Dim intEleveSource As Integer
    Dim intEleveDest As Integer

    ' *** FERMETURE USERFORM ***
    unloadAllUserForms
    
    ' *** AFFECTATION VARIABLES ***
    intClasseSource = 1 + listboxSelectionClasseSource.ListIndex
    strNomClasseSource = listboxSelectionClasseSource.Value
    intEleveSource = listboxSelectionEleve.ListIndex + 1
    
    strNomCompletEleve = Worksheets(strPage2).Cells(intLigListePage2 + intEleveSource, 2 * intClasseSource - 1)
    
    intClasseDest = 1 + listboxSelectionClasseDest.ListIndex
    strNomClasseDest = listboxSelectionClasseDest.Value
    intEleveDest = getIndiceEleve(strNomCompletEleve, intClasseDest, False)

    ' *** CONFIRMATION TRANSFERT ***
    If vbYes = MsgBox("Vous allez transférer '" & strNomCompletEleve & "' de la classe de '" & strNomClasseSource & "' vers la classe '" & strNomClasseDest & "'. Voulez-vous poursuivre ?", vbYesNo, "Confirmation de transfert") Then
        transfererEleve intClasseSource, intEleveSource, intClasseDest, intEleveDest
        MsgBox "Élève transféré."
    Else
        MsgBox "Operation annulée."
    End If
End Sub



