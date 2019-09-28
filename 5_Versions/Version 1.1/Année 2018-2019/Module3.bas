Attribute VB_Name = "Module3"
Sub Formules_Eleve()
'
' Formatage_Eleve Macro
'
' Touche de raccourci du clavier: Ctrl+Shift+F
'
    Dim Decal As Integer, Nb_Eleve As Integer, Col_deb As Integer, Col_fin As Integer
    Set ref = Worksheets("ref")
    
    Decal = ref.Range("P3").Value
    Col_T1 = ref.Range("E3").Value
    Col_T2 = ref.Range("F3").Value
    Col_T3 = ref.Range("G3").Value
    Col_An = ref.Range("H3").Value
    Eval_T1 = ref.Range("N3").Value
    Eval_T2 = ref.Range("N4").Value
    Eval_T3 = ref.Range("N5").Value
    Col = 3
    
    
    For i = 1 To Worksheets.Count
        If ActiveSheet.Name Like "Elève*" Then
            
            ' Formatage du 1er puis copie sur 2e élève
            '
            For Lig = 1 To 15
                If ref.Range(Cells(Lig, 10)).Value = "D*" Then
                    Lig_Domaine = ref.Cells(Lig, 11).Value
                    For Col = 1 To 3 + Decal
                        If Worksheets(1).Range(Cells(3, Col)).Value = "1er*trimestre" Then
                            =SIERREUR((((NB.SI('Elève (5ème1)'!C5:C9;"C"))*2)+((NB.SI('Elève (5ème1)'!C5:C9;"D"))*1))*$B$4/(NBVAL(C5:C9)+(((NB.SI('Elève (5ème1)'!C11:C12;"A"))*4)+((NB.SI('Elève (5ème1)'!C11:C12;"B"))*3)+((NB.SI('Elève (5ème1)'!C11:C12;"C"))*2)+((NB.SI('Elève (5ème1)'!C11:C12;"D"))*1))*$B$10/(NBVAL(C11:C12)+(((NB.SI('Elève (5ème1)'!C14:C15;"A"))*4)+((NB.SI('Elève (5ème1)'!C14:C15;"B"))*3)+((NB.SI('Elève (5ème1)'!C14:C15;"C"))*2)+((NB.SI('Elève (5ème1)'!C14:C15;"D"))*1))*$B$13/(NBVAL(C14:C15)+(((NB.SI('Elève (5ème1)'!C17:C19;"A"))*4)+((NB.SI('Elève (5ème1)'!C17:C19;"B"))*3)+((NB.SI('Elève (5ème1)'!C17:C19;"C"))*2)+((NB.SI('Elève (5ème1)'!C17:C19;"D"))*1))*$B$16/(NBVAL(C17:C19)+(((NB.SI('Elève (5ème1)'!C21:C21;"A"))*4)+((NB.SI('Elève (5ème1)'!C21:C21;"B"))*3)+((NB.SI('Elève (5ème1)'!C21:C21;"C"))*2)+((NB.SI('Elève (5ème1)'!C21:C21;"D"))*1))*$B$20/(NBVAL(C21:C21));"")
                            ((NB.SI('Elève (5ème1)'!C5:C9;"A"))*4)+
                            ((NB.SI('Elève (5ème1)'!C5:C9;"B"))*3)+
                            ((NB.SI('Elève (5ème1)'!C5:C9;"B"))*2)+
                            ((NB.SI('Elève (5ème1)'!C5:C9;"B"))*1)+
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            Cells(Lig, Col).Formula = "=SIERREUR((((NB.SI(Range(Cells(Col-1, Lig);""A""))*4)+((NB.SI('Elève (5ème1)'!C5:E9;""B""))*3)+((NB.SI('Elève (5ème1)'!C5:E9;""C""))*2)+((NB.SI('Elève (5ème1)'!C5:E9;""D""))*1))/(NBVAL(C5:E9));"""")"
                        End If
                    Next Col
                End If
            Next Lig
                
                    
            ' Visuel
            For Indice_eleve = 1 To 2 Step 1
                Range(Cells(2, Col), Cells(2, Col + Decal - 1)).HorizontalAlignment = xlCenterAcrossSelection
                Range(Cells(3, Col_T1 - 1), Cells(3, Col_T1)).HorizontalAlignment = xlCenterAcrossSelection
                Range(Cells(3, Col_T2 - 1), Cells(3, Col_T2)).HorizontalAlignment = xlCenterAcrossSelection
                Range(Cells(3, Col_T3 - 1), Cells(3, Col_T3)).HorizontalAlignment = xlCenterAcrossSelection
                Range(Cells(3, Col_An - 1), Cells(3, Col_An)).HorizontalAlignment = xlCenterAcrossSelection
                Union(Range(Cells(1, Col), Cells(2, Col)), Cells(4, Col), Cells(10, Col), Cells(13, Col), Cells(16, Col), Cells(20, Col), Range(Cells(1, Col_T1 + 1), Cells(2, Col_T1 + 1)), Cells(4, Col_T1 + 1), Cells(10, Col_T1 + 1), Cells(13, Col_T1 + 1), Cells(16, Col_T1 + 1), Cells(20, Col_T1 + 1), Range(Cells(1, Col_T2 + 1), Cells(2, Col_T2 + 1)), Cells(4, Col_T2 + 1), Cells(10, Col_T2 + 1), Cells(13, Col_T2 + 1), Cells(16, Col_T2 + 1), Cells(20, Col_T2 + 1)).Locked = True
                Union(Range(Cells(1, Col_T1 - 1), Cells(1, Col_T1)), Range(Cells(1, Col_T2 - 1), Cells(1, Col_T2)), Range(Cells(1, Col_T3 - 1), Cells(1, Col_T3 + 2))).EntireColumn.Locked = True
                
                Col = Col + Decal
                Col_T1 = Col_T1 + Decal
                Col_T2 = Col_T2 + Decal
                Col_T3 = Col_T3 + Decal
                Col_An = Col_An + Decal
            Next Indice_eleve
                
            
            
            'Formatage de tous les élèves à partir du modèle du 2e
            Appliquer_Tous_Eleves
            
            Set Plage = Nothing
        End If
  
End Sub

