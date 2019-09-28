Attribute VB_Name = "Module1"
' Procédure:
' Demande si Ajout ou Suppression
' Demande du trimestre
' Encas de suppression, autorise ou non
' Si autorisation, execution de la demande
' Sinon, fin de la macro

Sub Ajout_ou_Suppression_Eval()
Attribute Ajout_ou_Suppression_Eval.VB_ProcData.VB_Invoke_Func = "P\n14"
' Procédure globale d'ajout ou de suppression d'une évaluation
' Raccourci clavier: Crtl + Shift + P

    Dim ajout As Boolean, suppr As Boolean
    
    If ActiveSheet.Name Like "Elève*" Then
        If MsgBox("Souhaitez-vous ajouter une évaluation ?", vbYesNo, "Procédure d'ajout ou de suppression d'une évaluation") = vbYes Then
            ajout = True
        ElseIf MsgBox("Souhaitez-vous supprimer une évaluation ?", vbYesNo, "Procédure d'ajout ou de suppression d'une évaluation") = vbYes Then
            suppr = True
        Else
            If MsgBox("Aucune action effectuée.", vbOKOnly, "Procédure d'ajout ou de suppression d'une évaluation") = vbOK Then
            End If
        End If
        
        If ajout = True Or suppr = True Then
            num_tri = -1
            Do Until num_tri = 0 Or num_tri = 1 Or num_tri = 2 Or num_tri = 3
                num_tri = InputBox("Pour quel trimestre ? (1, 2 ou 3)" & Chr(10) & "Taper 0 pour annuler", "Choix du trimestre")
                If num_tri = vcCancel Then
                    num_tri = 0
                End If
            Loop
            If num_tri = 0 Then
            ElseIf num_tri = 1 Or num_tri = 2 Or num_tri = 3 Then
                If ajout = True Then
                    'MsgBox ("Ajout d'une évaluation au trimestre n°" & num_tri)
                    Ajout_Evaluation (num_tri)
                ElseIf suppr = True Then
                    If Autorisation_Suppression_Eval(num_tri) = True Then
                        'MsgBox ("Suppression d'une évaluation au trimestre n°" & num_tri)
                        Supprimer_Evaluation (num_tri)
                    Else: MsgBox ("Vous ne pouvez supprimer d'évaluation sur ce trimestre.")
                    End If
                End If
            Else: MsgBox ("Le numéro du semestre n'est pas valide, veuillez recommencer.")
            End If
        End If
    Else: MsgBox ("Veuillez vous placer sur une feuille 'Elève' pour éxecuter cette macro.")
    End If
            
End Sub

Sub Ajout_Evaluation(num_tri)
Attribute Ajout_Evaluation.VB_ProcData.VB_Invoke_Func = "A\n14"
' Amélioration : pouvoir executer la commande peu importe l'élève sélectionné
' Touche de raccourci du clavier: Ctrl+Shift+A
'
    Dim Decal As Integer, Colonne As Integer, Col_depart As Integer, Nb_Ele As Integer
    
    With Worksheets("ref")
        Decal = .Range("P3").Value
        Tri1 = .Range("N3").Value
        Tri2 = .Range("N4").Value
        Tri3 = .Range("N5").Value
    End With
    
    Select Case num_tri
        Case Is = "1"
            Col_depart = 2 + Tri1
        Case Is = "2"
            Col_depart = 4 + Tri1 + Tri2
        Case Is = "3"
            Col_depart = 6 + Tri1 + Tri2 + Tri3
    End Select
    Feuille_depart = ActiveSheet.Index
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name Like "Elève*" Then
            Worksheets(i).Select
            'MsgBox "Decalage: " & Decal
            
            Nb_Ele = Nb_Eleve()
            
            For Indice_eleve = 1 To Nb_Ele
                Colonne = Col_depart + (Indice_eleve - 1) * (Decal + 1)
                Cells(3, Colonne).EntireColumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                  
                Range(Cells(3, Colonne + 1), Cells(22, Colonne + 1)).Copy
                Range(Cells(3, Colonne), Cells(22, Colonne)).PasteSpecial Paste:=xlPasteValues
                Range(Cells(3, Colonne), Cells(22, Colonne)).PasteSpecial Paste:=xlPasteFormulas
                            
                If Indice_eleve = 1 Then
                    Cells(3, Colonne + 1).Copy
                    Cells(3, Colonne).PasteSpecial Paste:=xlPasteValues
                    Range(Cells(3, Colonne + 1), Cells(21, Colonne + 1)).ClearContents
                Else
                    Cells(3, Colonne).FormulaR1C1 = "=IF(RC[" & -Decal - 1 & "]="""","""",RC[" & -Decal - 1 & "])"
                    Range(Cells(4, Colonne + 1), Cells(21, Colonne + 1)).ClearContents
                End If
                
                Cells(3, Col_depart + 1).Select
                
            Next Indice_eleve
        End If
    Next i
    
    Worksheets(Feuille_depart).Select
    Cells(3, Col_depart + 1).Select

' Calcul du nombre d'évaluations par trimestre
    Nb_eval
End Sub

Function Autorisation_Suppression_Eval(num_tri)
' Renvoie un booléen qui indique si l'utilisateur est autorisé à supprimer une évaluation au trimestre donné"

    If Worksheets("ref").Range("N" & num_tri + 2).Value < 4 Then
        Autorisation_Suppression_Eval = False
    Else
        Autorisation_Suppression_Eval = True
    End If

End Function

Sub Supprimer_Evaluation(num_tri)
Attribute Supprimer_Evaluation.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' Touche de raccourci clavier: Crtl + Shift + S
'
    Dim Decal As Integer, Nb_Ele As Integer
    
    ' Récupération décalage & nb d'évaluation pour chaque trimestre
    With Worksheets("ref")
        Decal = .Range("P3").Value
        Tri1 = .Range("N3").Value
        Tri2 = .Range("N4").Value
        Tri3 = .Range("N5").Value
    End With
    
    Select Case num_tri
        Case Is = "1"
            Col_depart = 2 + Tri1
        Case Is = "2"
            Col_depart = 4 + Tri1 + Tri2
        Case Is = "3"
            Col_depart = 6 + Tri1 + Tri2 + Tri3
    End Select
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name Like "Elève*" Then
            Worksheets(i).Select
            
            Nb_Ele = Nb_Eleve()
            
            For Indice_eleve = 1 To Nb_Ele
                Colonne = Col_depart + (Indice_eleve - 1) * (Decal - 1)
                'Range(Cells(3, Colonne - 1), Cells(21, Colonne - 1)).Copy
                Cells(3, Colonne).EntireColumn.Delete Shift:=xlToLeft
                'Format_Eval (Colonne - 2)
            Next Indice_eleve
            
            Cells(3, Col_depart - 1).Select
            
        End If
    Next i
    
' Calcul du nombre d'évaluations par trimestre
    Nb_eval

End Sub

Sub Appliquer_Tous_Eleves()
Attribute Appliquer_Tous_Eleves.VB_ProcData.VB_Invoke_Func = "T\n14"
' Permet d'appliquer la sélection à tous les élèves (à partie de l'élève n°2)
' Raccourci clavier: Crtl + Shift + T
'
    Dim Decal As Integer, Nb_Lig As Integer, Lig_depart As Integer, Nb_El As Integer
    Decal = Worksheets("ref").Range("P3").Value
    Nb_Lig = Worksheets("ref").Range("K8").Value
        
    Nb_El = Nb_Eleve()
    
    Eleve_depart = 2
    Col_depart = 3 + (Eleve_depart - 1) * Decal
    
    For Indice_eleve = Eleve_depart To Nb_El
        If Indice_eleve = Eleve_depart Then
            Range(Cells(1, Col_depart), Cells(Nb_Lig, Col_depart + Decal - 1)).Copy
        Else
            New_col = Col_depart + (Indice_eleve - Eleve_depart) * Decal
            With Range(Cells(1, New_col), Cells(Nb_Lig, New_col + Decal - 1))
                .PasteSpecial Paste:=xlPasteFormats
                .PasteSpecial Paste:=xlPasteFormulas
            End With
        End If
    Next Indice_eleve
            
End Sub

Function Nb_Eleve()
' Calcul le nombre d'élève dans un tableur Elève
    Dim Classe As Worksheet
    
    If Range("A1").Value Like "Classe*" Then
        Set Classe = ActiveSheet
    Else
        Set Classe = Worksheets(Range("B2").Value)
    End If
    
    Lig = 4
    While Classe.Cells(Lig, 1).Value <> ""
        Lig = Lig + 1
    Wend
    
    Set Classe = Nothing
    Nb_Eleve = Lig - 4
    
End Function

Sub Nb_eval()
' Calcul du nombre d'évaluations par trimestre
' Touche de raccourci clavier: Crtl + Shift + E
    Dim Nb1 As Integer, Nb2 As Integer, Nb3 As Integer
    
    For Colonne = 3 To 100 Step 1
        If Cells(3, Colonne).Value Like "1er*trimestre" Then
            Nb1 = Colonne
        ElseIf Cells(3, Colonne).Value Like "2ème*trimestre" Then
            Nb2 = Colonne
        ElseIf Cells(3, Colonne).Value Like "3ème*trimestre" Then
            Nb3 = Colonne
            Exit For
        End If
    Next Colonne
    
    
    Worksheets("ref").Range("N3").Value = Nb1 - 3
    Worksheets("ref").Range("N4").Value = Nb2 - Nb1 - 2
    Worksheets("ref").Range("N5").Value = Nb3 - Nb2 - 2
    
End Sub

Sub Format_Eval(Col As Integer)
' Applique une bordure noire fine à toute la colonne sélectionnée
' Pas de raccourci clavier
           
    With Union(Range(Cells(3, Col), Cells(3, Col + 1)), Range(Cells(5, Col), Cells(9, Col + 1)), Range(Cells(11, Col), Cells(12, Col + 1)), Range(Cells(14, Col), Cells(15, Col + 1)), Range(Cells(17, Col), Cells(19, Col + 1)), Range(Cells(21, Col), Cells(21, Col + 1)))
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
    End With
    
    With Union(Cells(4, Col + 1), Cells(10, Col + 1), Cells(13, Col + 1), Cells(16, Col + 1), Cells(20, Col + 1))
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    With Range(Cells(1, Col), Cells(21, Col))
        .Locked = False
        .FormulaHidden = True
    End With
End Sub
