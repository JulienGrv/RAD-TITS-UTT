VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    MsgBox "R.A.D. est une macro de r�solution par analyse dimensionnelle.", vbOKOnly + vbInformation, "Bienvenue sur R.A.D. !"
    MsgBox "1) Compl�tez vos bases de donn�es (variables et constantes) en n'oubliant pas d'indiquer dans chaque table le symbole de la grandeur recherch�e." & Chr(10) & "2) Revenez sur la page principale 'analyse dimensionnelle' pour ajouter dans les colonnes respectives 'Variables' ou 'Constantes' le symbole de chaque grandeur que vous avez rentr� dans vos bases de donn�es." & Chr(10) & "3) Cliquez sur 'Ex�cuter la macro' et laissez vous guider !" & Chr(10) & "4) En cliquant sur 'Initialiser' vous effacerez tout ce qui se trouve sur la page principale : 'Analyse dimensionnelle'." & Chr(10) & "5) Initialisez la page 'Analyse dimensionnelle' d�s que vous changez de probl�me." & Chr(10) & "6) Vous pouvez revoir ces explications en cliquant sur 'Explication du programme'.", vbOKOnly + vbInformation, "Quelques explications"
End Sub
