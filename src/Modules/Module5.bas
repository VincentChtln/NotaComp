Attribute VB_Name = "Module5"

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
'               GNU General Public License V3 - Traduction française
'
'   Ce fichier fait partie de NotaComp.
'
'   NotaComp est un logiciel libre ; vous pouvez le redistribuer ou le modifier
'   suivant les termes de la GNU General Public License telle que publiée par la
'   Free Software Foundation, soit la version 3 de la Licence, soit (à votre gré)
'   toute version ultérieure.
'
'   NotaComp est distribué dans l’espoir qu’il sera utile, mais SANS AUCUNE
'   GARANTIE : sans même la garantie implicite de COMMERCIALISABILITÉ
'   ni d’ADÉQUATION À UN OBJECTIF PARTICULIER. Consultez la GNU
'   General Public License pour plus de détails.
'
'   Vous devriez avoir reçu une copie de la GNU General Public License avec NotaComp;
'   si ce n’est pas le cas, consultez : <http://www.gnu.org/licenses/>.
'
' *******************************************************************************


' *******************************************************************************
'                           Module 5 - Gestion des modules
' *******************************************************************************
'
'   Fonctions publiques
'
'   Procédures publiques
'
'   Fonctions privées
'       IsWBOpen(ByVal strWBName As String) As Boolean
'       IsVBProjectProtected(ByRef wb As Workbook) As Boolean
'       IsFolderEmpty(ByVal strFolderPath As String) As Boolean
'       ProperFolderPath(ByVal strFolderPath As String) As String
'       ProperWBName(ByVal strFileName As String) As String
'       GetDestFolder() As String
'
'   Procédures privées
'       DisplayUserName()
'       DisplayVBFilesInRecursiveFolder()
'       DeleteFile(ByVal strFileFullPath As String)
'       DeleteVBFilesInRecursiveFolder(ByVal strFolderPath As String)
'       UpdateVBProject()
'       ExportVBCode()
'       DeleteVBCode(ByRef wbTarget As Workbook)
'       ImportVBCode(ByRef wbTarget As Workbook)
'
' *******************************************************************************

Option Explicit

' *******************************************************************************
'                                   Constantes
' *******************************************************************************

Const strGitFolder      As String = "D:\Documents\GitHub\NotaComp\"
Const strWbSource       As String = "NotaComp_Dev.xlsm"

' *******************************************************************************
'                               Fonctions publiques
' *******************************************************************************

' *******************************************************************************
'                               Procédures publiques
' *******************************************************************************

' *******************************************************************************
'                               Fonctions privées
' *******************************************************************************

Private Function IsWBOpen(ByVal strWbName As String) As Boolean
    Dim wb As Workbook
    IsWBOpen = False
    For Each wb In Workbooks
        If wb.Name = strWbName Then
            IsWBOpen = True
            Exit For
        End If
    Next wb
End Function

Private Function IsVBProjectProtected(ByRef wb As Workbook) As Boolean
    IsVBProjectProtected = (wb.VBProject.Protection = vbext_pp_locked)
End Function

Private Function IsFolderEmpty(ByVal strFolderPath As String) As Boolean
    IsFolderEmpty = (Dir(ProperFolderPath(strFolderPath) & "*.*") = vbNullString)
End Function

Private Function ProperFolderPath(ByVal strFolderPath As String) As String
    ProperFolderPath = strFolderPath
    If Right$(ProperFolderPath, 1) <> "\" Then ProperFolderPath = ProperFolderPath & "\"
End Function

Private Function ProperWBName(ByVal strFileName As String) As String
    ProperWBName = strFileName
    If Right$(ProperWBName, 5) <> ".xlsm" Then ProperWBName = ProperWBName & ".xlsm"
End Function

Private Function GetDestFolder() As String
    GetDestFolder = strGitFolder
End Function

Private Function GetVBFilesInRecursiveFolder(ByVal strFolderPath As String) As String()
    ' *** VARIABLES ***
    Dim strFile         As String
    Dim strFullPath     As String
    Dim lFile           As Long
    Dim lNumFiles       As Long
    Dim arrFiles()      As String
    Dim arrNewFiles()   As String
    Dim lFolder         As Long
    Dim lNumFolders     As Long
    Dim arrFolders()    As String
    
    ' *** AFFECTATION VALEURS ***
    lNumFiles = 0
    lNumFolders = 0
    strFolderPath = ProperFolderPath(strFolderPath)
    strFile = Dir(strFolderPath & "*.*", vbDirectory)
    
    ' *** COMPTAGE FICHIERS + STOCKAGE DOSSIER DANS ARRAY ***
    While Len(strFile) <> 0
        If Left$(strFile, 1) <> "." Then
            strFullPath = strFolderPath & strFile
            
            ' *** AJOUT DANS ARRAY DOSSIERS ***
            If (GetAttr(strFullPath) And vbDirectory) = vbDirectory Then
                ReDim Preserve arrFolders(0 To lNumFolders) As String
                arrFolders(lNumFolders) = strFullPath
                lNumFolders = lNumFolders + 1
            
            ' *** AJOUT FICHIERS DANS ARRAY VBA ***
            Else
                If (Right$(strFile, 4) = ".cls" Or Right$(strFile, 4) = ".bas" Or Right$(strFile, 4) = ".frm" Or Right$(strFile, 4) = ".frx") Then
                    ReDim Preserve arrFiles(0 To lNumFiles) As String
                    arrFiles(lNumFiles) = strFullPath
                    lNumFiles = lNumFiles + 1
                End If
            End If
        End If
        strFile = Dir()
    Wend
    
    ' *** BOUCLE SUR ARRAY DE DOSSIERS ***
    For lFolder = 0 To lNumFolders - 1
        arrNewFiles = GetVBFilesInRecursiveFolder(arrFolders(lFolder))
        If GetSizeOfArray(arrFiles) <> 0 Then
            arrFiles = Split(Join(arrFiles, Chr(1)) & Chr(1) & Join(arrNewFiles, Chr(1)), Chr(1))
        Else
            arrFiles = arrNewFiles
        End If
    Next lFolder
    
    ' *** RENVOI VALEUR ***
    GetVBFilesInRecursiveFolder = arrFiles
            
End Function

' *******************************************************************************
'                               Procédures privées
' *******************************************************************************

'@EntryPoint
Public Sub DisplayUserName()
    MsgBox "Current user is '" & Application.UserName & "'."
End Sub

Public Sub DisplayVBFilesInRecursiveFolder()
    ' *** VARIABLES ***
    Dim strFile         As Variant
    Dim arrFiles()      As String
    Dim strListFiles    As String
    
    ' *** RECUPERATION ARRAY FICHIERS ***
    arrFiles = GetVBFilesInRecursiveFolder(GetDestFolder() & "src\")
    
    ' *** CONSTRUCTION ET AFFICHAGE LISTE ***
    For Each strFile In arrFiles
        strListFiles = strListFiles & "  - " & strFile & vbNewLine
    Next strFile
    MsgBox ("Liste fichiers VBA : " & vbNewLine & strListFiles)
End Sub

Private Sub DeleteFile(ByVal strFileFullPath As String)
    If Dir(strFileFullPath) <> vbNullString Then
        SetAttr strFileFullPath, vbNormal
        Kill strFileFullPath
    End If
End Sub

Private Sub DeleteVBFilesInRecursiveFolder(ByVal strFolderPath As String)
    Dim arrFiles() As String
    Dim varFile As Variant
    
    strFolderPath = ProperFolderPath(strFolderPath)
    arrFiles = GetVBFilesInRecursiveFolder(strFolderPath)
    
    If GetSizeOfArray(arrFiles) <> 0 Then
        For Each varFile In arrFiles
            On Error Resume Next
            Kill varFile
            On Error GoTo 0
        Next varFile
    End If
End Sub

'@EntryPoint
Private Sub UpdateVBProject()
    ' *** VARIABLES ***
    Dim wbTarget As Workbook
    Dim bWbOK As Boolean
    Dim bWbOpen As Boolean
    Dim strWbName As String
    Dim strWbFolder As String
    
    DisableUpdates
    
    ' *** EXPORT NOUVEAU CODE ***
    If (vbYes = MsgBox("Suppression ancien code et export nouveau code ?", vbYesNo)) Then ExportVBCode
    
    ' *** CREATION NOUVEAU CLASSEUR NOTACOMP ***
    If (vbYes = MsgBox("Création nouveau Workbook 'NotaComp' ?", vbYesNo)) Then
        Set wbTarget = Workbooks.Add
        strWbName = "NotaComp"
        strWbFolder = ProperFolderPath(GetDestFolder())
        Call ImportVBCode(wbTarget)
        Call DeleteFile(strWbFolder & strWbName & ".xlsm")
        wbTarget.SaveAs Filename:=strWbFolder & strWbName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Call Application.Run("'NotaComp.xlsm'!InitNotaComp")
        wbTarget.Save
        wbTarget.Close
        GoTo EOS
    End If
    
    ' *** MAJ 1 CLASSEUR UNIQUEMENT ***
    If (vbYes = MsgBox("MAJ classeur existant ?", vbYesNo)) Then
        bWbOK = False
        bWbOpen = False
        While Not bWbOK
            strWbName = ProperWBName(InputBox("Nom du classeur cible ?", "Nom du classeur", strWbName))
            If IsWBOpen(strWbName) Then
                bWbOK = True
                bWbOpen = True
            Else
                strWbFolder = ProperFolderPath(InputBox("Chemin vers le classeur cible", "Chemin du dossier", strWbFolder))
                If Dir(strWbFolder & strWbName) <> vbNullString Then
                    bWbOK = True
                    Call Workbooks.Open(strWbFolder & strWbName)
                    While Not IsWBOpen(strWbName)
                    Wend
                Else
                    If vbCancel = MsgBox("Fichier non trouvé, vérifier l'orthographe", vbOKCancel) Then Exit Sub
                End If
            End If
        Wend
        Set wbTarget = Workbooks(strWbName)
        If IsVBProjectProtected(wbTarget) Then
            Call MsgBox("Projet VB bloqué, impossible d'importer le nouveau code.", vbInformation)
        Else
            Call DeleteVBCode(wbTarget)
            Call ImportVBCode(wbTarget)
            wbTarget.Save
            If Not bWbOpen Then wbTarget.Close
        End If
        GoTo EOS
    End If
    
    ' *** MAJ TOUS LES CLASSEURS D'UN DOSSIER ***
    If (vbYes = MsgBox("MAJ plusieurs classeurs existants situés dans le même dossier ?", vbYesNo)) Then
        
    End If
    
EOS:
    EnableUpdates
End Sub

Private Sub ExportVBCode()
    ' *** CONSTANTES ***
    Const Padding = 24
    
    ' *** VARIABLES ***
    Dim VBComp              As VBIDE.VBComponent
    Dim byNbExportedFiles   As Byte
    Dim strPath             As String
    Dim strFolder           As String
    Dim strSubFolder        As String
    Dim strExtension        As String
    Dim bFolderOK           As Boolean
    
    ' *** VERIFICATION ACCES VBPROJECT ***
    If IsVBProjectProtected(Workbooks(strWbSource)) Then
        Call MsgBox("Projet VB protégé, accès refusé.")
        Exit Sub
    End If
    
    ' *** SUPPRESSION ANCIENS FICHIERS ***
    strFolder = ProperFolderPath(ThisWorkbook.Path)
    Call DeleteVBFilesInRecursiveFolder(strFolder)
    
    ' *** AJOUT DOSSIERS SI NECESSAIRE ***
    If Dir(strFolder & "Documents", vbDirectory) = vbNullString Then MkDir (strFolder & "Documents")
    If Dir(strFolder & "UserForms", vbDirectory) = vbNullString Then MkDir (strFolder & "UserForms")
    If Dir(strFolder & "Modules", vbDirectory) = vbNullString Then MkDir (strFolder & "Modules")
    
    ' *** EXPORT NOUVEAU CODE ***
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        If VBComp.CodeModule.CountOfLines > 0 Then
            Select Case VBComp.Type
                Case vbext_ct_Document
                    strSubFolder = "\Documents"
                    strExtension = ".cls"
                Case vbext_ct_MSForm
                    strSubFolder = "\UserForms"
                    strExtension = ".frm"
                Case vbext_ct_StdModule
                    strSubFolder = "\Modules"
                    strExtension = ".bas"
                Case Else
                    GoTo NextIteration
            End Select
            
            On Error Resume Next
            Err.Clear
            
            strPath = strFolder & strSubFolder & "\" & VBComp.Name & strExtension
            Call VBComp.Export(strPath)
            
            If Err.Number <> 0 Then
                Call MsgBox("Error #" & Str(Err.Number) & ". Failed to export " & VBComp.Name & " to " & strPath, vbCritical)
            Else
                byNbExportedFiles = byNbExportedFiles + 1
                Debug.Print "Exported " & Left$(VBComp.Name & ":" & Space(Padding), Padding) & strPath
            End If
    
            On Error GoTo 0
        End If
NextIteration:
    Next VBComp
    
    Call DisplayTemporaryMessage("Export réussi de " & CStr(byNbExportedFiles) & " fichiers VBA vers " & strFolder, 10)
End Sub

Private Sub DeleteVBCode(ByRef wbTarget As Workbook)
    ' *** DECLARATION VARIABLES ***
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    ' *** AFFECTATION VARIABLES ***
    Set VBProj = wbTarget.VBProject
    
    ' *** SUPPRESSION FICHIERS VBA ***
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type <> vbext_ct_Document Then
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
    
    ' *** MESSAGE INFORMATION ***
    Call DisplayTemporaryMessage("Modules et UserForm VBA supprimés.", 10)
End Sub

Private Sub ImportVBCode(ByRef wbTarget As Workbook)
    Dim strImportFolder     As String
    Dim bWbOK               As Boolean
    Dim arrVBFiles()        As String
    Dim varFile             As Variant
    Dim byNbImportedFiles   As Byte
    Dim modSource           As Variant
    Dim modDest             As Variant
    Dim VBComp              As VBIDE.VBComponent

    If ThisWorkbook.Name = wbTarget.Name Then
        MsgBox ("Import impossible pour " & wbTarget.Name & " : classeur source")
        Exit Sub
    End If
    
    If Not IsWBOpen(wbTarget.Name) Then
        Call MsgBox("Import impossible pour " & wbTarget.Name & " : classeur non ouvert")
        Exit Sub
    End If
    
    If IsVBProjectProtected(wbTarget) Then
        Call MsgBox("Import impossible pour " & wbTarget.Name & " : projet VB protégé")
        Exit Sub
    End If
    
    Call DeleteVBCode(wbTarget)
    
    strImportFolder = ProperFolderPath(ThisWorkbook.Path)
    arrVBFiles = GetVBFilesInRecursiveFolder(strImportFolder)
    
    For Each varFile In arrVBFiles
        If (Right$(varFile, 4) = ".bas" Or Right$(varFile, 4) = ".frm") Then
            wbTarget.VBProject.VBComponents.Import varFile
            byNbImportedFiles = byNbImportedFiles + 1
        End If
    Next varFile
    
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        If VBComp.Type = vbext_ct_Document And VBComp.CodeModule.CountOfLines > 0 Then
            On Error Resume Next
            Set modSource = VBComp.CodeModule
            Set modDest = wbTarget.VBProject.VBComponents(VBComp.Name).CodeModule
            With modDest
                .DeleteLines StartLine:=1, Count:=.CountOfLines
                .AddFromString modSource.Lines(1, modSource.CountOfLines)
            End With
            byNbImportedFiles = byNbImportedFiles + 1
            On Error GoTo 0
        End If
    Next VBComp
    
    Call DisplayTemporaryMessage("Import réussi de " & CStr(byNbImportedFiles) & " fichiers VBA vers " & wbTarget.Name, 10)
End Sub
