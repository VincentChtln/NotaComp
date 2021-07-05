VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Modification classe"
   ClientHeight    =   2640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Procédure d'ajout d'un élève
Private Sub btnAjouterEleve_Click()
    UserForm2.Show
End Sub

' Procédure de suppression d'un élève
Private Sub btnSupprimerEleve_Click()

    ' *** MESSAGE ALERTE ***
    If vbCancel = MsgBox("ATTENTION: vous allez supprimer un élève d'une classe. Ses notes seront également perdues dans le processus." & vbNewLine & _
                         "Cette opération est irréversible. Si vous souhaitez toutefois garder ses notes, veuillez les enregister dans un document à part." & vbNewLine & vbNewLine & _
                         "Pour revenir à la fenêtre précédente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte") Then Exit Sub
    
    UserForm3.Show
End Sub

' Procédure de transfert d'un élève
Private Sub btnTransfererEleve_Click()

    ' *** MESSAGE ALERTE ***
    If vbCancel = MsgBox("ATTENTION: vous allez transférer un élève entre d'une classe vers une autre." & vbNewLine & _
                         "Puisque rien ne garantit la comptabilité des évaluations entre deux classes différentes, " & _
                         "les notes de l'élève ne seront pas transférées et seront par conséquent perdues dans le processus." & vbNewLine & _
                         "Si vous souhaitez les conserver, veillez à les relever dans un document à part. " & _
                         "Il tient ensuite à vous d'adapter les notes précédemment acquises aux évaluations de sa nouvelle classe." & vbNewLine & vbNewLine & _
                         "Si vous souhaitez revenir à la fenêtre précédente, cliquez sur 'Annuler'.", vbOKCancel, "Message d'alerte") Then Exit Sub
     
    UserForm4.Show
End Sub


