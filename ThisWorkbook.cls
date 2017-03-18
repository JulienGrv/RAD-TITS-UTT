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
    MsgBox "R.A.D. est une macro de résolution par analyse dimensionnelle.", vbOKOnly + vbInformation, "Bienvenue sur R.A.D. !"
    MsgBox "1) Complétez vos bases de données (variables et constantes) en n'oubliant pas d'indiquer dans chaque table le symbole de la grandeur recherchée." & Chr(10) & "2) Revenez sur la page principale 'analyse dimensionnelle' pour ajouter dans les colonnes respectives 'Variables' ou 'Constantes' le symbole de chaque grandeur que vous avez rentré dans vos bases de données." & Chr(10) & "3) Cliquez sur 'Exécuter la macro' et laissez vous guider !" & Chr(10) & "4) En cliquant sur 'Initialiser' vous effacerez tout ce qui se trouve sur la page principale : 'Analyse dimensionnelle'." & Chr(10) & "5) Initialisez la page 'Analyse dimensionnelle' dès que vous changez de problème." & Chr(10) & "6) Vous pouvez revoir ces explications en cliquant sur 'Explication du programme'.", vbOKOnly + vbInformation, "Quelques explications"
End Sub
