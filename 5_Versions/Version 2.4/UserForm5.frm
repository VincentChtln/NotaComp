VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Opération en cours"
   ClientHeight    =   1104
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Function limiterPourcentage(intValeur As Integer) As Integer
    limiterPourcentage = intValeur
    If limiterPourcentage < 0 Then
        limiterPourcentage = 0
    ElseIf limiterPourcentage > 100 Then
        limiterPourcentage = 100
    End If
End Function

Public Sub updateAvancement(intAvancementActuel As Integer, intAvancementTotal As Integer)
    Dim intPourc As Integer
    intPourc = Me.limiterPourcentage(100 * intAvancementActuel / intAvancementTotal)
    UserForm5.lblChargement.Caption = intPourc & "% terminé"
    UserForm5.barChargement.Width = 2 * intPourc
    UserForm5.Repaint
End Sub
