VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Supprimer un élève"
   ClientHeight    =   5136
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3780
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ##################################
' UF Suppression d'élève
' ##################################

Option Explicit

' ##################################
' PROCÉDURES
' ##################################
' UserForm_Initialize()
' listboxSelectionClasse_Change()
' btnSupprimerEleve_Click()
' ##################################

' Initialisation de l'UF
Private Sub UserForm_Initialize()
    ' *** DECLARATION VARIABLES ***
    Dim intNbClasse As Integer
    Dim intIndiceClasse As Integer
    Dim strNomClasse As String
    
    ' *** AFFECTATION VARIABLES ***
    intNbClasse = getNombreClasses()
    
    ' *** AJOUT CLASSE DANS LISTE ***
    For intIndiceClasse = 1 To intNbClasse
        strNomClasse = getNomClasse(intIndiceClasse)
        listboxSelectionClasse.AddItem strNomClasse
    Next intIndiceClasse
    
    ' *** INITIALISATION SELECTION ***
    listboxSelectionClasse.ListIndex = 0
    listboxSelectionEleve.ListIndex = 0
End Sub

' Modification de la liste Eleve en fonction de la classe sélectionnée
Private Sub listboxSelectionClasse_Change()
    ' *** DECLARATION VARIABLES ***
    Dim intIndiceClasse As Integer
    Dim intColonneClasse As Integer
    Dim intNbEleves As Integer
    Dim intIndiceEleve As Integer
    Dim strNomCompletEleve As String
    
    ' *** AFFECTATION VARIABLES ***
    listboxSelectionEleve.Clear
    intIndiceClasse = listboxSelectionClasse.ListIndex + 1
    intColonneClasse = 2 * intIndiceClasse - 1
    intNbEleves = getNombreEleves(intIndiceClasse)
    
    ' *** AJOUT ELEVES DANS LISTE ***
    For intIndiceEleve = 1 To intNbEleves
        strNomCompletEleve = Worksheets(strPage2).Cells(intLigListePage2 + intIndiceEleve, intColonneClasse).Value
        listboxSelectionEleve.AddItem strNomCompletEleve
    Next intIndiceEleve
    
    ' *** INITIALISATION INDEX ***
    listboxSelectionEleve.ListIndex = 0
End Sub

' Demande de confirmation puis appel de la procédure supprimerEleve (Module 2)
Private Sub btnSupprimerEleve_Click()
    Dim intIndiceClasse As Integer
    Dim strNomClasse As String
    Dim intIndiceEleve As Integer
    Dim strNomComplet As String
    
    intIndiceClasse = 1 + listboxSelectionClasse.ListIndex
    strNomClasse = listboxSelectionClasse.Value
    intIndiceEleve = listboxSelectionEleve.ListIndex + 1
    strNomComplet = Worksheets(strPage2).Cells(intLigListePage2 + intIndiceEleve, 2 * intIndiceClasse - 1)
    
    If vbYes = MsgBox("Vous êtes sur le point de supprimer '" & strNomComplet & "' de la classe de " & strNomClasse & ". Voulez-vous poursuivre ?", vbYesNo, "Confirmation de suppression") Then
        supprimerEleve intIndiceClasse, intIndiceEleve
        MsgBox "Élève supprimé"
    Else
        MsgBox "Operation annulée"
    End If
End Sub


