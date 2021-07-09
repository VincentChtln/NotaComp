VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Transférer un élève "
   ClientHeight    =   5040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' *******************************************************************************
' UF Transfert d'élève
' *******************************************************************************

Option Explicit

' *******************************************************************************
' PROCÉDURES
' *******************************************************************************
' UserForm_Initialize()
' lbxClasseSource_Change()
' BtnTransfererEleve_Click()
' *******************************************************************************

' Initialisation de l'UF
' Peuplement de lbxClasseSource
Private Sub UserForm_Initialize()
    ' *** DECLARATION VARIABLES ***
    Dim byNbClasses As Byte
    Dim byClasse As Byte
    
    ' *** AFFECTATION VARIABLES ***
    byNbClasses = GetNombreClasses
    
    ' *** AJOUT CLASSES DANS LISTE ***
    For byClasse = 1 To byNbClasses
        lbxClasseSource.AddItem GetNomClasse(byClasse)
    Next byClasse
    
    ' *** INITIALISATION INDEX ***
    lbxClasseSource.ListIndex = 0
End Sub

' Modification de la liste Eleve en fonction de la classe sélectionnée
Private Sub lbxClasseSource_Change()
    ' *** DECLARATION VARIABLES ***
    Dim byEleve As Byte
    Dim byNbEleves As Byte
    Dim byClasseSource As Byte
    Dim byClasseDest As Byte
    Dim byNbClasses As Byte
    Dim byColClasseSource As Byte

    ' *** AFFECTATION VARIABLES ***
    lbxEleve.Clear
    lbxClasseDest.Clear
    byClasseSource = lbxClasseSource.ListIndex + 1
    byColClasseSource = 2 * byClasseSource - 1
    byNbEleves = GetNombreEleves(byClasseSource)
    byNbClasses = GetNombreClasses

    ' *** AJOUT ELEVES DANS LISTE ***
    For byEleve = 1 To byNbEleves
        lbxEleve.AddItem ThisWorkbook.Worksheets(strPage2).Cells(byLigListePage2 + byEleve, byColClasseSource).Value
    Next byEleve
    
    ' *** AJOUT CLASSES DANS LISTE TRANSFERT ***
    For byClasseDest = 1 To byNbClasses
        If byClasseDest <> byClasseSource Then
            lbxClasseDest.AddItem GetNomClasse(byClasseDest)
        End If
    Next byClasseDest

    ' *** INITIALISATION INDEX ***
    lbxEleve.ListIndex = 0
    lbxClasseDest.ListIndex = 0
End Sub

' Demande de confirmation puis appel de la procédure transfererEleve (Module 2)
Private Sub BtnTransfererEleve_Click()
    ' *** DECLARATION VARIABLES ***
    Dim byClasseSource As Byte
    Dim strClasseSource As String
    Dim byClasseDest As Byte
    Dim strClasseDest As String
    Dim strEleve As String
    Dim byEleveSource As Byte
    Dim byEleveDest As Byte

    ' *** FERMETURE USERFORM ***
    UnloadAllUserForms
    
    ' *** AFFECTATION VARIABLES ***
    byEleveSource = lbxEleve.ListIndex + 1
    byClasseSource = lbxClasseSource.ListIndex + 1
    byClasseDest = lbxClasseDest.ListIndex + 1
    strEleve = ThisWorkbook.Worksheets(strPage2).Cells(byLigListePage2 + byEleveSource, 2 * byClasseSource - 1)
    strClasseSource = lbxClasseSource.Value
    strClasseDest = lbxClasseDest.Value
    byEleveDest = GetIndiceEleve(strEleve, byClasseDest, False)

    ' *** CONFIRMATION TRANSFERT ***
    If vbYes = MsgBox("Vous allez transférer '" & strEleve & "' de la classe '" & strClasseSource & "' vers la classe '" & strClasseDest & "'. " & _
                      "Confirmez-vous cette opération ?", vbYesNo, "Confirmation de transfert") Then
        TransfererEleve byClasseSource, byEleveSource, byClasseDest, byEleveDest, strEleve
        MsgBox "Élève transféré."
    Else
        MsgBox "Operation annulée."
    End If
End Sub


