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

' **********************************
' PROCÉDURES
' **********************************
' UserForm_Initialize()
' listboxSelectionClasseSource_Change()
' btnTransfererEleve_Click()
' **********************************

' Initialisation de l'UF
' Peuplement de listboxSelectionClasseSource
Private Sub UserForm_Initialize()
    Dim intNombreClasse As Integer, intIndiceClasse As Integer, strNomClasse As String
    
    intNombreClasse = getNombreClasses()
    
    For intIndiceClasse = 1 To intNombreClasse
        strNomClasse = getNomClasse(intIndiceClasse)
        listboxSelectionClasseSource.AddItem strNomClasse
    Next intIndiceClasse
    
    listboxSelectionClasseSource.ListIndex = 0
End Sub

' Modification de la liste Eleve en fonction de la classe sélectionnée
Private Sub listboxSelectionClasseSource_Change()

    ' Peuplement de listboxSelectionEleve
    Dim intIndiceClasseSource As Integer, intColonneClasseSource As Integer
    Dim intNombreEleves As Integer, intIndiceEleve As Integer, strNomCompletEleve As String

    listboxSelectionEleve.Clear
    intIndiceClasseSource = listboxSelectionClasseSource.ListIndex + 1
    intColonneClasseSource = 2 * intIndiceClasseSource - 1
    intNombreEleves = getNombreEleves(intIndiceClasseSource)

    For intIndiceEleve = 1 To intNombreEleves
        strNomCompletEleve = Cells(3 + intIndiceEleve, intColonneClasseSource).Value
        listboxSelectionEleve.AddItem strNomCompletEleve
    Next intIndiceEleve

    listboxSelectionEleve.ListIndex = 0
    
    ' Peuplement de listboxSelectionClasseDest
    Dim intNombreClasses As Integer
    Dim intIndiceClasseDest As Integer, strNomClasseDest As String

    listboxSelectionClasseDest.Clear

    intNombreClasses = getNombreClasses()

    For intIndiceClasseDest = 1 To intNombreClasses
        If intIndiceClasseDest <> intIndiceClasseSource Then
            strNomClasseDest = getNomClasse(intIndiceClasseDest)
            listboxSelectionClasseDest.AddItem strNomClasseDest
        End If
    Next intIndiceClasseDest

    listboxSelectionClasseDest.ListIndex = 0
    
End Sub

' Demande de confirmation puis appel de la procédure transfererEleve (Module 2)
Private Sub btnTransfererEleve_Click()
    Dim intClasseSource As Integer, strNomClasseSource As String
    Dim intClasseDest As Integer, strNomClasseDest As String
    Dim strNomCompletEleve As String, intEleveSource As Integer, intEleveDest As Integer
    Dim obj As Object

    ' Fermeture des UF
    unloadAllUserForms
    
    ' Données classe source
    intClasseSource = 1 + listboxSelectionClasseSource.ListIndex
    strNomClasseSource = listboxSelectionClasseSource.Value
    intEleveSource = listboxSelectionEleve.ListIndex + 1

    strNomCompletEleve = Cells(3 + intEleveSource, 2 * intClasseSource - 1)

    ' Données classe dest
    intClasseDest = 1 + listboxSelectionClasseDest.ListIndex
    strNomClasseDest = listboxSelectionClasseDest.Value
    intEleveDest = getIndiceEleve(strNomCompletEleve, intClasseDest, False)

    ' Confirmation de transfert
    If vbYes = MsgBox("Vous allez transférer '" & strNomCompletEleve & "' de la classe de '" & strNomClasseSource & "' vers la classe '" & strNomClasseDest & "'. Voulez-vous poursuivre ?", vbYesNo, "Confirmation de transfert") Then
'        transfererEleve intClasseSource, intEleveSource, intClasseDest, intEleveDest
        MsgBox "Élève transféré."
    Else
        MsgBox "Operation annulée."
    End If
End Sub



