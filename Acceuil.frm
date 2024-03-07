VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Acceuil 
   Caption         =   "GESTIONS PRESENCE DES ETUDIANTS"
   ClientHeight    =   8880.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11340
   OleObjectBlob   =   "Acceuil.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Acceuil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ligne_debut As Integer: Dim colone_debut As Integer
Dim ligneFin As Integer: Dim coloneFin As Integer
Dim ligne_enCours As Integer: Dim colone_enCours As Integer

Private Sub affichage_liste_Click()

End Sub
Private Sub ajouter_seance_Click()

End Sub

Private Sub Btnimporter_fichier_Click()
    'btn importation fichier'
    Dim fichier_importer As String
    
    fichier_importer = Application.GetOpenFilename("Fichier excel(*.xlsx),*.xlsx", , "selectionner le fichier excel a importer")
End Sub

Private Sub ChoixPrensence_Change()
' Obtenir le choix de l'utilisateur
    Dim choix As String
    choix = ChoixPrensence.Value

    ' Exécuter une action en fonction du choix de l'utilisateur
    Select Case choix
        Case "Présent"
            ' Code à exécuter si l'utilisateur choisit "Présent"
            Selection.Style = "Satisfaisant"
            ActiveCell.FormulaR1C1 = "Present"
            Range("E2").Select
        Case "Absent"
            ' Code à exécuter si l'utilisateur choisit "Absent"
            Selection.Style = "Insatisfaisant"
            ActiveCell.FormulaR1C1 = "Absent"
            Range("E2").Select
        Case "Excusé"
            ' Code à exécuter si l'utilisateur choisit "Justifié"
            Selection.Style = "Neutre"
            ActiveCell.FormulaR1C1 = "Excusé"
            Range("E2").Select
    End Select
End Sub

Private Sub UserForm_Initialize()
    ' Ajouter les choix à la ComboBox
    With ChoixPrensence
        .AddItem "Présent"
        .AddItem "Absent"
        .AddItem "Excusé"
    End With
End Sub

Private Sub filtrer_selon_matricule_Change()

End Sub

Private Sub Statistique_etudiant_Click()

End Sub

Private Sub UserForm_Click()
    
End Sub
