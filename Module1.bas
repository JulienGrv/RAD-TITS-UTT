Attribute VB_Name = "Module1"
'Une grandeur est une variable ayant une dimension
Private Type grandeur
    variable As Integer
    dimension(6) As Double '6 car il y a 7 dimensions de base
    ordreGrandeur As Double
End Type

'Ma macro
Private Sub PhysiqueSansEquation()
    
    Worksheets("Analyse dimensionnelle").Activate
    On Error GoTo NotValidInput
    Range(Cells(7, 3), Cells(15, 255)).Select
    Selection.Clear
    Selection.Columns.ColumnWidth = 10.71
    
    'Cells.Clear
    
    'On détermine le nombre de variables
    Dim nbVariables As Integer
    Dim choixEtude As Integer
    choixEtude = 0
    
    While (choixEtude <> 1 And choixEtude <> 2)
        choixEtude = InputBox("L'étude porte t-elle sur les variables ou les constantes ?" & Chr(10) & "Saisir 1 pour les variables ou 2 pour les constantes.")
    Wend
    
    nbVariables = 0
    Cells(2, choixEtude).Select
    While Not (IsEmpty(ActiveCell.Offset(nbVariables, 0)))
        nbVariables = nbVariables + 1
    Wend
    
    'On dimensionne notre tableau contenant nos grandeurs avec la taille nbVariables
    Dim tabGrandeur() As grandeur
    ReDim tabGrandeur(nbVariables - 1)
    
    'On enregistre nos grandeurs dans le tableau tabGrandeur
    Dim i As Integer
    Dim indGrandeur As Integer
    Dim indDimension As Integer
    Cells(2, choixEtude).Select
    If (choixEtude = 1) Then
        For indGrandeur = 0 To nbVariables - 1
            i = 3
            'On recherche la variable dans la base de données des grandeurs
            While Worksheets("Base de données des grandeurs").Cells(i, 2).Value <> ActiveCell.Value
                i = i + 1
                If IsEmpty(Worksheets("Base de données des grandeurs").Cells(i, 2)) Then
                    MsgErrBox1 ("Vous avez saisi une variable qui n'est pas présente dans la base de données !" & Chr(10) & "Allez dans la feuille base de données des grandeurs et ajoutez la grandeur.")
                    Exit Sub
                End If
            Wend
            tabGrandeur(indGrandeur).variable = i
            For indDimension = 0 To 6
                tabGrandeur(indGrandeur).dimension(indDimension) = Worksheets("Base de données des grandeurs").Cells(i, 4 + indDimension).Value
            Next
            tabGrandeur(indGrandeur).ordreGrandeur = Worksheets("Base de données des grandeurs").Cells(i, 11).Value
            ActiveCell.Offset(1, 0).Select
        Next
    ElseIf choixEtude = 2 Then
        For indGrandeur = 0 To nbVariables - 1
            i = 3
            'On recherche la variable dans la base de données des grandeurs
            While Worksheets("Base de données des constantes").Cells(i, 2).Value <> ActiveCell.Value
                i = i + 1
                If IsEmpty(Worksheets("Base de données des constantes").Cells(i, 2)) Then
                    MsgErrBox1 ("Vous avez saisi une constante qui n'est pas présente dans la base de données !" & Chr(10) & "Allez dans la feuille base de données des constantes et ajoutez la constante.")
                    Exit Sub
                End If
            Wend
            tabGrandeur(indGrandeur).variable = i
            For indDimension = 0 To 6
                tabGrandeur(indGrandeur).dimension(indDimension) = Worksheets("Base de données des constantes").Cells(i, 4 + indDimension).Value
            Next
            tabGrandeur(indGrandeur).ordreGrandeur = Worksheets("Base de données des constantes").Cells(i, 11).Value
            ActiveCell.Offset(1, 0).Select
        Next
    End If
    
    'On demande à l'utilisateur la variable à isoler
    Dim variableIsolee As Integer
    variableIsolee = 0
    variableIsolee = InputBox("Entrer la ligne de la variable isolée :")
    If variableIsolee < 2 Or variableIsolee > nbVariables + 1 Then
        MsgErrBox1 ("La ligne doit être comprise entre 2 et " & nbVariables + 1)
    End If
    
    'On détermine le nombre de combinaisons de fonctions à tester
    Dim nbCombinaisons As Integer
    Dim k As Integer
    Dim n As Integer
    k = 0
    k = InputBox("k parmi n" & Chr(10) & "Entrer k (nombre de dimensions fondamentales du problème) :")
    If k < 2 Or k > nbVariables - 1 Then
        MsgErrBox1 ("k doit être compris entre 2 et " & nbVariables - 1)
    End If
    n = 0
    n = InputBox("k parmi n" & Chr(10) & "Entrer n (nombres de variables sans prendre en compte la variable isolée) :" & Chr(10) & "Exemple : t,v,m,g,h => n=4")
    If n <> nbVariables - 1 Then
        MsgErrBox1 ("n doit valoir " & nbVariables - 1 & Chr(10) & "Vous n'avez pas saisi le bon nombre de variables !")
    End If
    nbCombinaisons = Factorielle(n) / (Factorielle(k) * Factorielle(n - k))
    
    'On initialise les paramètres qui vont servir à créer les combinaisons
    Dim j As Integer
    Dim compteur As Integer
    Dim indCombinaison As Integer
    Dim choix As Integer
    Dim IsPresent As Boolean
    Dim combinaison() As Integer
    ReDim combinaison(nbCombinaisons - 1, k - 1)
        
    Cells(7, 3).Value = "Fonctions :"
    Cells(7, 4).Select
    Randomize

    'On créer les combinaisons
    For indCombinaison = 0 To nbCombinaisons - 1
        If choixEtude = 1 Then
            ActiveCell.Value = Worksheets("Base de données des grandeurs").Cells(tabGrandeur(variableIsolee - 2).variable, 2).Value & "=f("
        Else
            ActiveCell.Value = Worksheets("Base de données des constantes").Cells(tabGrandeur(variableIsolee - 2).variable, 2).Value & "=f("
        End If
        'On initialise la combinaison
        For i = 0 To k - 1
            combinaison(indCombinaison, i) = -1
        Next
        
        'On créer une combinaison
        For indGrandeur = 0 To k - 1
            Do
                choix = Int(nbVariables * Rnd)
            Loop While choix = variableIsolee - 2
            IsPresent = False
            For i = 0 To indGrandeur
                If combinaison(indCombinaison, i) = choix Then
                    IsPresent = True
                    Exit For
                End If
            Next
            If Not (IsPresent) Then
                combinaison(indCombinaison, indGrandeur) = choix
            Else
                indGrandeur = indGrandeur - 1
            End If
        Next
        
        'On tri la combinaison (tri bulle)
        For i = 0 To k - 1
            For j = i To k - 1
                If combinaison(indCombinaison, i) > combinaison(indCombinaison, j) Then
                    Temp = combinaison(indCombinaison, j)
                    combinaison(indCombinaison, j) = combinaison(indCombinaison, i)
                    combinaison(indCombinaison, i) = Temp
                End If
            Next
        Next
        
        'On vérifie que cette combinaison n'existe pas déjà
        IsPresent = False
        If indCombinaison <> 0 Then
            For i = 0 To indCombinaison - 1
                compteur = 0
                For indGrandeur = 0 To k - 1
                    If combinaison(i, indGrandeur) = combinaison(indCombinaison, indGrandeur) Then
                        compteur = compteur + 1
                    End If
                Next
                If compteur = k Then
                    IsPresent = True
                    Exit For
                End If
            Next
        End If
        
        'On l'affiche ou pas
        If Not (IsPresent) Then
            'On écrit la combinaison
            For indGrandeur = 0 To k - 1
                If indGrandeur <> 0 Then
                    ActiveCell.Value = ActiveCell.Value & ","
                End If
                If choixEtude = 1 Then
                    ActiveCell.Value = ActiveCell.Value & Worksheets("Base de données des grandeurs").Cells(tabGrandeur(combinaison(indCombinaison, indGrandeur)).variable, 2).Value
                Else
                    ActiveCell.Value = ActiveCell.Value & Worksheets("Base de données des constantes").Cells(tabGrandeur(combinaison(indCombinaison, indGrandeur)).variable, 2).Value
                End If
            Next
            ActiveCell.Value = ActiveCell.Value & ")"
            ActiveCell.Offset(0, 2).Select
        Else
            indCombinaison = indCombinaison - 1
        End If
    Next
    
    'Inscrire la dimension de la variable isolée puissance a1
    Dim colonne As Integer
    Cells(8, 4).Select
    For indCombinaison = 0 To nbCombinaisons - 1
        ActiveCell.Value = DimensionText(tabGrandeur(variableIsolee - 2).variable, choixEtude) & "=("
        ActiveCell.Offset(0, 2).Select
    Next
    
    'Inscrit les dimensions avec leur exposant associé
    Cells(8, 4).Select
    For indCombinaison = 0 To nbCombinaisons - 1
        For indGrandeur = 0 To k - 1
            If indGrandeur <> 0 Then
                ActiveCell.Value = ActiveCell.Value & "*("
            End If
            ActiveCell.Value = ActiveCell.Value & DimensionText(tabGrandeur(combinaison(indCombinaison, indGrandeur)).variable, choixEtude) & ")^a" & indGrandeur + 1
        Next
        ActiveCell.Offset(0, 2).Select
    Next
    
    'AX=Y => X=(A^-1)Y
    Dim A1() As Double
    Dim X() As Double
    ReDim X(k - 1)
    Dim Y() As Double
    Dim nbZero As Integer
    Dim error As Boolean
    Cells(10, 5).Select
    
    'On résoud toutes les équations
    For indCombinaison = 0 To nbCombinaisons - 1
        'On créée A et Y avec 7 lignes car on ignore encore les dimensions fondamentales qui entrent en jeu
        ReDim A1(k - 1, 6)
        ReDim Y(6)
        'On remplis A avec les dimensions des variables et Y avec la dimension de la variable isolée
        For j = 0 To k - 1
            For i = 0 To 6
                A1(j, i) = tabGrandeur(combinaison(indCombinaison, j)).dimension(i)
                Y(i) = tabGrandeur(variableIsolee - 2).dimension(i)
            Next
        Next
        'On cherche les lignes de A qui sont toutes à zéro et on les supprime car cela signifie que la dimension relative à la ligne n'intervient pas dans l'analyse dimensionnelle
        For i = 0 To 6
            nbZero = 0
            'On regarde si la ligne a toutes ses valeurs à zéro
            For j = 0 To k - 1
                If A1(j, i) = 0 Then
                    nbZero = nbZero + 1
                End If
            Next
            'Partie où on inverse les lignes avant de diminuer la taille de la matrice A
            If nbZero = k And i <> 6 Then
                IsPresent = False
                For choix = i + 1 To 6
                    For j = 0 To k - 1
                        If A1(j, choix) <> 0 Then
                            IsPresent = True
                            Exit For
                        End If
                    Next
                    If IsPresent = True Then
                        Exit For
                    End If
                Next
                If IsPresent Then
                    For j = 0 To k - 1
                        A1(j, i) = A1(j, choix)
                        A1(j, choix) = 0
                    Next
                    Y(i) = Y(choix)
                    Y(choix) = 0
                Else
                    ReDim Preserve A1(k - 1, i - 1)
                    ReDim Preserve Y(i - 1)
                    Exit For
                End If
            ElseIf nbZero = k And i = 6 Then
                ReDim Preserve A1(k - 1, i - 1)
                ReDim Preserve Y(i - 1)
                Exit For
            End If
        Next
        'On vérifie que la matrice est carrée
        If UBound(A1, 1) <> UBound(A1, 2) Then
            ActiveCell.Offset(-1, -1).Value = "pas de solutions réelles : la matrice n'est pas carrée !"
            ActiveCell.Offset(k, 0).Select
        Else
            'On transpose la matrice
            Dim A2() As Double
            colonne = UBound(A1, 1)
            ligne = UBound(A1, 2)
            ReDim A2(ligne, colonne)
            A2 = Transposee(A1)
            'On l'inverse avec la méthode du pivot de Gauss (très rapide)
            error = False
            A2 = InverseMatrice(A2, error)
            If Not (error) Then
                'On calcul les solutions
                X = CalculSolution(A2, Y)
                'On les affiche
                ActiveCell.Offset(-1, -1).Value = "exposants"
                ActiveCell.Offset(-1, 0).Value = "valeurs"
                For i = 0 To k - 1
                    ActiveCell.Offset(0, -1).Value = "a" & i + 1 & " ="
                    ActiveCell.Value = X(i)
                    ActiveCell.Offset(1, 0).Select
                Next
                Dim resultat As Double
                resultat = 1
                For i = 0 To k - 1
                    resultat = resultat * Application.WorksheetFunction.Power(tabGrandeur(combinaison(indCombinaison, i)).ordreGrandeur, X(i))
                Next
                resultat = Application.WorksheetFunction.Round(resultat, 2)
                If choiEtude = 1 Then
                    ActiveCell.Offset(0, -1).Value = "application numérique " & Worksheets("Base de données des grandeurs").Cells(tabGrandeur(variableIsolee - 2).variable, 2).Value & "=" & resultat & " "
                Else
                    ActiveCell.Offset(0, -1).Value = "application numérique " & Worksheets("Base de données des constantes").Cells(tabGrandeur(variableIsolee - 2).variable, 2).Value & "=" & resultat & " "
                End If
                For i = 0 To 6
                    If tabGrandeur(variableIsolee - 2).dimension(i) <> 0 Then
                    If choixEtude = 1 Then
                        ActiveCell.Offset(0, -1).Value = ActiveCell.Offset(0, -1).Value & ConvDimToSI(Worksheets("Base de données des grandeurs").Cells(2, i + 4).Value) & "^" & tabGrandeur(variableIsolee - 2).dimension(i)
                    Else
                        ActiveCell.Offset(0, -1).Value = ActiveCell.Offset(0, -1).Value & ConvDimToSI(Worksheets("Base de données des constantes").Cells(2, i + 4).Value) & "^" & tabGrandeur(variableIsolee - 2).dimension(i)
                    End If
                    End If
                Next
            Else
                ActiveCell.Offset(-1, -1).Value = "pas de solutions réelles : le déterminant est nul !"
                For i = 0 To k - 1
                    ActiveCell.Offset(1, 0).Select
                Next
            End If
        End If
        ActiveCell.Offset(-k, 2).Select
    Next
    
    Range(Cells(1, 1), Cells(40, 2 * nbCombinaisons)).Columns.AutoFit
    MsgBox ("Terminé !")
    End
    
NotValidInput:
    MsgErrBox1 ("Vous avez entrer une valeur invalide (Type mismatch) !")
    
End Sub

'Affichage des erreurs
Private Sub MsgErrBox1(ByVal Message As String)
    MsgBox Message, vbCritical, "PhysiqueSansEquation"
    End
End Sub
Private Sub MsgErrBox2(ByVal Message As String)
    MsgBox Message, vbCritical, "Factorielle"
    End
End Sub
Private Sub MsgErrBox3(ByVal Message As String)
    MsgBox Message, vbCritical, "DimensionText"
    End
End Sub
Private Sub MsgErrBox4(ByVal Message As String)
    MsgBox Message, vbCritical, "inverseMatrice"
    End
End Sub

'La fonction factorielle permet de calculer la factorielle d'un nombre entier n
Private Function Factorielle(n As Integer) As Double
    If n = 0 Or n = 1 Then
        Factorielle = 1
    ElseIf n > 1 And n < 171 Then
        Factorielle = 1
        Dim i As Integer
        For i = 2 To n
            Factorielle = Factorielle * i
        Next
    Else
        MsgErrBox2 ("Veuillez saisir un nombre entier compris entre 0 et 170 !")
    End If
End Function

'La fonction dimension convertit la valeur dimensionnelle d'une variable sous forme d'une chaine de caractère
Private Function DimensionText(ligne As Integer, choixEtude As Integer) As String
    Dim i As Integer
    Dim numerateur As String
    Dim denominateur As String
    numerateur = ""
    denominateur = ""
    If choixEtude = 1 Then
        For i = 0 To 6
            If Worksheets("Base de données des grandeurs").Cells(ligne, 4 + i).Value = 1 Then
                numerateur = numerateur & Worksheets("Base de données des grandeurs").Cells(2, 4 + i).Value
            ElseIf Worksheets("Base de données des grandeurs").Cells(ligne, 4 + i).Value = -1 Then
                denominateur = denominateur & Worksheets("Base de données des grandeurs").Cells(2, 4 + i).Value
            ElseIf Worksheets("Base de données des grandeurs").Cells(ligne, 4 + i).Value > 0 Then
                numerateur = numerateur & Worksheets("Base de données des grandeurs").Cells(2, 4 + i).Value & "^" & Abs(Worksheets("Base de données des grandeurs").Cells(ligne, 4 + i).Value)
            ElseIf Worksheets("Base de données des grandeurs").Cells(ligne, 4 + i).Value < 0 Then
                denominateur = denominateur & Worksheets("Base de données des grandeurs").Cells(2, 4 + i).Value & "^" & Abs(Worksheets("Base de données des grandeurs").Cells(ligne, 4 + i).Value)
            End If
        Next
    Else
        For i = 0 To 6
            If Worksheets("Base de données des constantes").Cells(ligne, 4 + i).Value = 1 Then
                numerateur = numerateur & Worksheets("Base de données des constantes").Cells(2, 4 + i).Value
            ElseIf Worksheets("Base de données des constantes").Cells(ligne, 4 + i).Value = -1 Then
                denominateur = denominateur & Worksheets("Base de données des constantes").Cells(2, 4 + i).Value
            ElseIf Worksheets("Base de données des constantes").Cells(ligne, 4 + i).Value > 0 Then
                numerateur = numerateur & Worksheets("Base de données des constantes").Cells(2, 4 + i).Value & "^" & Abs(Worksheets("Base de données des constantes").Cells(ligne, 4 + i).Value)
            ElseIf Worksheets("Base de données des constantes").Cells(ligne, 4 + i).Value < 0 Then
                denominateur = denominateur & Worksheets("Base de données des constantes").Cells(2, 4 + i).Value & "^" & Abs(Worksheets("Base de données des constantes").Cells(ligne, 4 + i).Value)
            End If
        Next
    End If
    If denominateur <> "" And numerateur <> "" Then
        DimensionText = numerateur & "/" & denominateur
    ElseIf denominateur = "" Then
        DimensionText = numerateur
    ElseIf numerateur = "" Then
        DimensionText = "1/" & denominateur
    Else
        If choixEtude = 1 Then
            MsgErrBox3 ("La variable : " & Worksheets("Base de données des grandeurs").Cells(ligne, 2).Value & " n'a aucune dimension !")
        Else
            MsgErrBox3 ("La variable : " & Worksheets("Base de données des constantes").Cells(ligne, 2).Value & " n'a aucune dimension !")
        End If
    End If
End Function

'La fonction inverse la matrice avec la méthode de Gauss (source : http://codes-sources.commentcamarche.net/source/23266-inversion-de-matrices)
Private Function InverseMatrice(ByRef matrice() As Double, ByRef error As Boolean) As Double()
    Dim i As Integer, j As Integer, k As Integer, jmax As Integer
    Dim n As Integer
    Dim M() As Double, MInv() As Double
    Dim Temp As Double, Max As Double

    n = UBound(matrice, 1)
    
    ' vérifie que la matrice est une matrice carrée
    If UBound(matrice, 2) <> n Then MsgErrBox4 ("La matrice n'est pas carrée !")
    
    ' crée la matrice n x 2n, composée par M et la matrice identité
    ReDim M(n, 2 * n + 1)
    For i = 0 To n
        For j = 0 To n
            M(i, j) = matrice(i, j)
            M(i, j + n + 1) = 1 - Sgn(Abs(i - j))
        Next
    Next
    
    ' échelonne la matrice M()
    For i = 0 To n
        ' trouve le pivot maximum
        j = i
        Max = 0
        For k = j To n
            If Abs(M(k, i)) > Max Then
                jmax = k
                Max = Abs(M(k, i))
            End If
        Next
        If Max = 0 Then
            error = True
            GoTo NotInversibleMatrice
        End If
        
        j = jmax
        ' échange les 2 lignes si elles sont différentes
        ' commence à partir de l'élément i, car tous les précédents sont nuls
        If i <> j Then
            For k = i To 2 * n + 1
                Temp = M(i, k)
                M(i, k) = M(j, k)
                M(j, k) = Temp
            Next
        End If
        ' le pivot devient égal à 1
        If M(i, i) <> 1 Then
            Temp = M(i, i)
            For j = i To 2 * n + 1
                M(i, j) = M(i, j) / Temp
            Next
        End If
        ' sous le pivot, tous les éléments deviennent nuls
        For j = i + 1 To n
            If M(j, i) <> 0 Then
                Temp = M(j, i)
                For k = i To 2 * n + 1
                    M(j, k) = M(j, k) - M(i, k) * Temp
                Next
            End If
        Next
    Next
    
    ' réduit la matrice M()
    For i = n To 1 Step -1
        For j = 0 To i - 1
            If M(j, i) <> 0 Then
                Temp = M(j, i)
                For k = i To 2 * n + 1
                    M(j, k) = M(j, k) - M(i, k) * Temp
                Next
            End If
        Next
    Next
        
    ' retourne le résultat : la deuxième partie de la matrice M()
    ReDim MInv(n, n)
    For i = 0 To n
        For j = 0 To n
            MInv(i, j) = M(i, j + n + 1)
        Next
    Next
    
    InverseMatrice = MInv
NotInversibleMatrice:
End Function

'La fonction retourne la transposee d'une matrice
Private Function Transposee(ByRef matrice() As Double) As Double()
    Dim mTransposee() As Double
    colonne = UBound(matrice, 1)
    ligne = UBound(matrice, 2)
    Dim i As Integer
    Dim j As Integer
    ReDim mTransposee(ligne, colonne)
    
    For i = 0 To ligne
        For j = 0 To colonne
            mTransposee(i, j) = matrice(j, i)
        Next
    Next
    
    Transposee = mTransposee
End Function

'La fonction retourne le produit matriciel de A par B
Private Function CalculSolution(ByRef A() As Double, ByRef B() As Double) As Double()
    Dim produit() As Double
    ligne = UBound(A, 1)
    Dim i As Integer
    Dim j As Integer
    ReDim produit(ligne)
    
    For i = 0 To ligne
        produit(i) = ProduitLC(A, B, i)
    Next
    
    CalculSolution = produit
End Function

'La fonction retourne le produit d'une ligne par une colonne
Private Function ProduitLC(ByRef A() As Double, ByRef B() As Double, ligne As Integer) As Double
    Dim k As Integer
    ProduitLC = 0
    
    For k = 0 To UBound(A, 2)
        ProduitLC = ProduitLC + A(ligne, k) * B(k)
    Next
    
End Function

'Convertisseur dimension => unite SI
Private Function ConvDimToSI(dimension As String) As String
    If dimension = "L" Then
        ConvDimToSI = "m"
    ElseIf dimension = "M" Then
        ConvDimToSI = "kg"
    ElseIf dimension = "T" Then
        ConvDimToSI = "s"
    ElseIf dimension = "I" Then
        ConvDimToSI = "A"
    ElseIf dimension = Worksheets("Base de données des grandeurs").Cells(2, 8).Value Then
        ConvDimToSI = "K"
    ElseIf dimension = "J" Then
        ConvDimToSI = "Cd"
    ElseIf dimension = "N" Then
        ConvDimToSI = "mol"
    End If
End Function

Private Sub InitAnalyse()
    
    Worksheets("Analyse dimensionnelle").Activate
    
    Range(Cells(7, 3), Cells(15, 255)).Select
    Selection.Clear
    Selection.Columns.ColumnWidth = 10.71
    
    Range(Cells(2, 1), Cells(200, 2)).Select
    Selection.Clear
    Selection.Columns.ColumnWidth = 10.71
    
End Sub

Private Sub Explication()
    MsgBox "R.A.D. est une macro de résolution par analyse dimensionnelle." & Chr(10) & "", vbOKOnly + vbInformation, "Bienvenue sur R.A.D. !"
    MsgBox "1) Complétez vos bases de données (variables et constantes) en n'oubliant pas d'indiquer dans chaque table le symbole de la grandeur recherchée." & Chr(10) & "2) Revenez sur la page principale 'analyse dimensionnelle' pour ajouter dans les colonnes respectives 'Variables' ou 'Constantes' le symbole de chaque grandeur que vous avez rentré dans vos bases de données." & Chr(10) & "3) Cliquez sur 'Exécuter la macro' et laissez vous guider !" & Chr(10) & "4) En cliquant sur 'Initialiser' vous effacerez tout ce qui se trouve sur la page principale : 'Analyse dimensionnelle'." & Chr(10) & "5) Initialisez la page 'Analyse dimensionnelle' dès que vous changez de problème." & Chr(10) & "6) Vous pouvez revoir ces explications en cliquant sur 'Explication du programme'.", vbOKOnly + vbInformation, "Quelques explications"
End Sub

