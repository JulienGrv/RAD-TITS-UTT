Attribute VB_Name = "Module3"
Private Sub AjouterConstante()
    
    Worksheets("Base de données des constantes").Activate
    ActiveSheet.Unprotect ("jujuseb")
    On Error GoTo NotValidInput
    
    Dim nbVariables As Integer
    Dim grandeur As String
    nbVariables = 0
    
    Cells(3, 2).Select
    While Not (IsEmpty(ActiveCell.Offset(nbVariables, 0)))
        nbVariables = nbVariables + 1
    Wend
    
    grandeur = ""
    grandeur = InputBox("Saisir la constante à ajouter :" & Chr(10) & "Syntaxe : variable(L,M,T,I,K,J,N)" & Chr(10) & Chr(10) & "Exemple : h(2,1,-1,0,0,0,0)")
    If grandeur = "" Then
        MsgBox "Vous devez saisir une constante !", vbCritical, "AjouterConstante"
        ActiveSheet.Protect ("jujuseb")
        End
     ElseIf NbOc(grandeur, "(") <> 1 Or NbOc(grandeur, ")") <> 1 Or NbOc(grandeur, ",") <> 6 Then
        MsgBox "Vous n'avez pas correctement saisi la constante. Respectez la syntaxe !", vbCritical, "AjouterConstante"
        ActiveSheet.Protect ("jujuseb")
        End
    End If
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim dimension As Double
    Dim ordreGrandeur As Double
    Dim description As String
    
    If nbVariables > 0 Then
        Cells(nbVariables + 3, 2).Select
        ActiveCell.EntireRow.Insert
        ActiveCell.Value = Left(grandeur, InStr(grandeur, "(") - 1)
        ActiveCell.Offset(0, 2).Select
        
        j = 1
        k = 1
        For i = 0 To 6
            If Mid(grandeur, InStr(grandeur, "(") + 2 * i + k, j) = "-" Then
                j = 2
            End If
            dimension = Mid(grandeur, InStr(grandeur, "(") + 2 * i + k, j)
            ActiveCell.Offset(0, i).Value = dimension
            If j = 2 Then
                k = k + 1
                j = 1
            End If
        Next
        
        description = ""
        While (description = "")
            description = InputBox("Saisir une description ou un commentaire personnel sur la constante :")
        Wend
        ActiveCell.Offset(0, -1).Value = description
        
        ordreGrandeur = InputBox("Saisir la valeur de la constante en unité S.I." & Chr(10) & "Info : Entrer 1 pour la grandeur recherchée." & Chr(10) & " 1,7x10^-12 s'écrira 1,7e-12")
        ActiveCell.Offset(0, 7).Value = ordreGrandeur
    ElseIf nbVariables = 0 Then
        Cells(3, 2).Select
        description = ""
        While (description = "")
            description = InputBox("Saisir une description ou un commentaire personnelle sur la constante :")
        Wend
        ActiveCell.Offset(0, 1).Value = description
        ActiveCell.Value = Left(grandeur, InStr(grandeur, "(") - 1)
        ActiveCell.Offset(0, 2).Select
        
        j = 1
        k = 1
        For i = 0 To 6
            If Mid(grandeur, InStr(grandeur, "(") + 2 * i + k, j) = "-" Then
                j = 2
            End If
            ActiveCell.Offset(0, i).Value = Mid(grandeur, InStr(grandeur, "(") + 2 * i + k, j)
            If j = 2 Then
                k = k + 1
                j = 1
            End If
        Next
        ordreGrandeur = InputBox("Saisir la valeur de la constante en unité S.I." & Chr(10) & "Info : 1,7x10^-12 s'écrira 1,7e-12")
        ActiveCell.Offset(0, 7).Value = ordreGrandeur
    End If
    
    Range(Cells(2, 2), Cells(nbVariables + 3, 11)).Select
    Selection.Columns.AutoFit
    Selection.Rows.AutoFit
    
    ResizeGrandeur (nbVariables)
    ActiveSheet.Protect ("jujuseb")
    End
    
NotValidInput:
    MsgBox "Vous avez entrer une valeur invalide (Type mismatch) !", vbCritical, "SupprimerConstante"
    ActiveCell.EntireRow.Delete
    ActiveSheet.Protect ("jujuseb")
End Sub

Private Sub SupprimerConstante()
    
    Worksheets("Base de données des constantes").Activate
    ActiveSheet.Unprotect ("jujuseb")
    On Error GoTo NotValidInput
    
    Dim ligneVariable As String
    
    ligneVariable = 0
    ligneVariable = InputBox("Saisir la ligne de la constante à supprimer :")
    If ligneVariable = 0 Then
        MsgBox "Vous devez saisir une ligne !", vbCritical, "SupprimerConstante"
        ActiveSheet.Protect ("jujuseb")
        End
    End If
    
    Cells(ligneVariable, 2).Select
    If ActiveCell.Value <> "" Then
        ActiveCell.EntireRow.Delete
    Else
        MsgBox "La ligne saisie est vide !", vbCritical, "SupprimerConstante"
    End If
    
    ResizeGrandeur (1000 + ligneVariable)
    ActiveSheet.Protect ("jujuseb")
    End
    
NotValidInput:
    MsgBox "Vous avez entrer une valeur invalide (Type mismatch) !", vbCritical, "SupprimerConstante"
    ActiveSheet.Protect ("jujuseb")
    
End Sub

Private Sub InitCte()
    
    Worksheets("Base de données des constantes").Activate
    ActiveSheet.Unprotect ("jujuseb")
    
    Dim nbVariables As Integer
    nbVariables = 0
    
    Cells(3, 2).Select
    While Not (IsEmpty(ActiveCell.Offset(nbVariables, 0)))
        nbVariables = nbVariables + 1
    Wend
    
    If nbVariables <> 0 Then
        Cells(nbVariables + 3, 2).Select
        ActiveCell.EntireRow.Insert
        
        Range(Cells(3, 2), Cells(nbVariables + 2, 11)).Select
        Selection.Delete Shift:=xlUp
    End If
    
    ResizeGrandeur (1000 + nbVariables)
    ActiveSheet.Protect ("jujuseb")
    
End Sub

Private Sub ModifierValeurConstante()
    
    Worksheets("Base de données des constantes").Activate
    ActiveSheet.Unprotect ("jujuseb")
    On Error GoTo NotValidInput
    
    Dim grandeur As Integer
    grandeur = InputBox("Saisir la ligne de la valeur de constante à ajouter/modifier :")
    
    Cells(grandeur, 11).Select
    If Not IsEmpty(ActiveCell.Offset(0, -1).Value) Then
        Dim ordreGrandeur As Double
        ordreGrandeur = InputBox("Saisir la nouvelle valeur en unité S.I." & Chr(10) & "Info : 1,7x10^-12 s'écrira 1,7e-12")
        ActiveCell.Value = ordreGrandeur
    End If
    
    ResizeGrandeur (1000 + grandeur)
    ActiveSheet.Protect ("jujuseb")
    End
    
NotValidInput:
    MsgBox "Vous avez entrer une valeur invalide (Type mismatch) !", vbCritical, "SupprimerConstante"
    ActiveSheet.Protect ("jujuseb")
    
End Sub

Private Sub ModifierDescriptionCte()
    
    Worksheets("Base de données des constantes").Activate
    ActiveSheet.Unprotect ("jujuseb")
    
    Dim grandeur As Integer
    grandeur = InputBox("Saisir la ligne de la description à modifier :")
    
    Cells(grandeur, 3).Select
    If Not IsEmpty(ActiveCell.Offset(0, -1).Value) Then
        Dim ordreGrandeur As String
        ordreGrandeur = InputBox("Saisir la nouvelle description :")
        ActiveCell.Value = ordreGrandeur
    End If
    
    ResizeGrandeur (1000 + grandeur)
    ActiveSheet.Protect ("jujuseb")
    
End Sub

Private Function ResizeGrandeur(ByVal taille As Integer)
    Worksheets("Base de données des constantes").Activate
    Range(Cells(2, 2), Cells(taille + 3, 11)).Select
    Selection.Columns.AutoFit
    Selection.Rows.AutoFit
End Function
