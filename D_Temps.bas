Attribute VB_Name = "D_Temps"
Sub GestionDesTemps()
    Debug.Print "Début de la gestion des temps de production"
    
    Configurateur.CalculTempsUtile 'Ce sub permet de calculer le temps utile disponible pour chaque centre
    Debug.Print "Temps utile calculé pour chaque centre"
    
    Initialisation_TempsCU 'Ce sub permet de lire les fichiers de paramètres des centres pour récupérer les capacités et cadences de profils
    Debug.Print "Initialisation des temps des centres d'usinage terminée"
    
    CalculCADENCES 'Calcul des cadences des centres
    Debug.Print "Calcul des cadences terminé"
    
    Debug.Print "Fin de la gestion des temps de production"
End Sub

Function GetControlValue(Form As Object, ControlName As String) As Variant
    On Error Resume Next
    GetControlValue = Configurateur.Controls(ControlName).Value
    On Error GoTo 0
End Function

Sub SetControlValue(Form As Object, ControlName As String, Value As Variant)
    On Error Resume Next
    Configurateur.Controls(ControlName).Caption = Value
    On Error GoTo 0
End Sub

Public Sub Initialisation_TempsCU()
    Debug.Print "Début de l'initialisation des temps des centres d'usinage"
    
    Dim j As Long
    Dim Fichiers As Variant
    Dim Contenus As Variant
    Dim Temps As Variant
    Dim ligne As Variant
    Dim Champs() As String
    Dim Fichier As Object
    
    'Chemin vers les fichiers de paramètres
    Fichiers = Array( _
        "U:\06- Projet\04- Alu\42 - Passage sur CU V14\00_Paramètres\01_Données Centres\DUBUS3.txt", _
        "U:\06- Projet\04- Alu\42 - Passage sur CU V14\00_Paramètres\01_Données Centres\DUBUS5.txt", _
        "U:\06- Projet\04- Alu\42 - Passage sur CU V14\00_Paramètres\01_Données Centres\DUBUS8.txt", _
        "U:\06- Projet\04- Alu\42 - Passage sur CU V14\00_Paramètres\01_Données Centres\SCHIRMER4.txt")
    
    Dim TempsDicts(4) As Object
    For j = 0 To 4
        Set TempsDicts(j) = CreateObject("Scripting.Dictionary")
    Next j
    
    ' Parcours de chaque fichier
    For j = 0 To UBound(Fichiers)
        Set Fichier = Manipulationfichiers.OpenTextFile(Fichiers(j), 1)
        Contenus = Split(Fichier.readall, vbNewLine)
        
        ' Parcours des lignes du fichier
        For Each ligne In Contenus
            Champs = Split(ligne, ";")
            If UBound(Champs) > 0 Then
                TempsDicts(j)(Champs(1)) = Champs(2)  ' Stockage dans le dictionnaire correspondant
            End If
        Next ligne
        Fichier.Close
    Next j
    
    ' Assignation des dictionnaires aux variables globales
    Set Temps_DUBUS3 = TempsDicts(0)
    Set Temps_DUBUS5 = TempsDicts(1)
    Set Temps_DUBUS8 = TempsDicts(2)
    Set Temps_SCHIRMER4 = TempsDicts(3)
    
    Debug.Print "Initialisation des temps terminée"
End Sub

Public Sub CalculCADENCES()
    Debug.Print "Début du calcul des cadences"
    
    Dim oDossierLotsCU As Object
    Dim oFichierLotsCU As Object
    Dim FichierLOTPACKtxt As Object
    Dim ContenuLOTPACKtxt() As String
    Dim ChampsLOTPACKtxt() As String
    Dim i As Integer
    Dim ProfilLOTPACK As String
    Dim QteLOTPACK As Integer
    Dim Charge As Double
    Dim Centres As Variant
    Dim DictionnaireKeys As Variant
    Dim centre As String
    Dim Nfichier As Long
    Dim Clef As String
    Dim NombreCentresCompatibles As Integer
    Dim CentreSelect As String
    
    ' Centres d'usinage et les clefs MES pour chaque centre
    Centres = Array("DUBUS3", "DUBUS5", "DUBUS8", "SCHIRMER4")
    DictionnaireKeys = Array("DUBUS3-2", "DUBUS5-2", "DUBUS8-2", "SCHIRM4-2")
    
    ' Récupération du dossier contenant les fichiers LOTPACK
    Set oDossierLotsCU = Manipulationfichiers.GetFolder("U:\06- Projet\04- Alu\42 - Passage sur CU V14\%TEMP%")
    
    ' Affichage de l'étape en cours sur l'interface utilisateur
    Dim Niter As Long
    Niter = 0  ' Initialisation du compteur de fichiers
    
    ' Boucle pour traiter chaque fichier dans le dossier
    For Each oFichierLotsCU In oDossierLotsCU.Files
        Niter = Niter + 1 ' Incrémentation du compteur de fichiers
        
        ' Déclaration et initialisation des tableaux pour chaque centre
        Dim Compatible(4) As Integer
        Dim Charges(4) As Double
        For j = 0 To 4
            Compatible(j) = 1 ' Initialiser la compatibilité à 1 (compatible)
            Charges(j) = 0 ' Initialiser la charge à 0
        Next j
        
        ' Ouverture du fichier et lecture de son contenu
        Set FichierLOTPACKtxt = Manipulationfichiers.OpenTextFile(oFichierLotsCU, 1)
        ContenuLOTPACKtxt = Split(FichierLOTPACKtxt.readall, vbNewLine)
        
        ' Boucle sur chaque ligne du fichier
        For i = 0 To UBound(ContenuLOTPACKtxt) - 1
            ' Séparation de la ligne en champs
            ChampsLOTPACKtxt = Split(ContenuLOTPACKtxt(i), ";")
            ProfilLOTPACK = Left(ChampsLOTPACKtxt(0), 6) ' Extraction du profil
            QteLOTPACK = ChampsLOTPACKtxt(1) ' Extraction de la quantité
            
            ' Boucle pour chaque centre d'usinage
            For j = 0 To UBound(Centres)
                centre = Centres(j)
                
                ' Si le centre est compatible (1), on calcule la charge
                If Compatible(j) = 1 Then
                    Charge = RécupérationCharge(centre, ProfilLOTPACK, QteLOTPACK)
                    If Charge >= 0 Then
                        Charges(j) = Charges(j) + Charge + PénalitésCU(ProfilLOTPACK, Charge, centre, Mid(oFichierLotsCU.Name, InStr(oFichierLotsCU.Name, "_") + 1, InStr(oFichierLotsCU.Name, ".") - InStr(oFichierLotsCU.Name, "_") - 1)) ' On ajoute la charge au total
                    Else
                        Compatible(j) = 0 ' Si incompatible, on marque comme non compatible (0)
                        Charges(j) = 0
                    End If
                End If
            Next j
        Next i
        
        ' Après avoir analysé toutes les lignes, on met à jour le dictionnaire avec les charges
        For j = 0 To UBound(Centres)
            ' Si une charge est positive pour ce centre, on l'enregistre dans le dictionnaire
            If Charges(j) <> 0 Then
                CentreSelect = DictionnaireKeys(j)
                LotSelect = Mid(oFichierLotsCU.Name, InStr(oFichierLotsCU.Name, "_") + 1, InStr(oFichierLotsCU.Name, ".") - InStr(oFichierLotsCU.Name, "_") - 1) & "$" & DictionnaireKeys(j)
                Dictionnaire_ChargeCU(LotSelect) = Charges(j)
            End If
        Next j
        
        NombreCentresCompatibles = 0
        For Each item In Charges
            If item <> 0 Then NombreCentresCompatibles = NombreCentresCompatibles + 1
        Next item
        
        If NombreCentresCompatibles = 0 Then Debug.Print "Aucun centre compatible pour ce lot."
        
        ' Mise à jour de la barre de progression tous les 10 fichiers traités
        If Niter Mod 10 = 0 Then Configurateur.Maj_Progression Niter, oDossierLotsCU.Files.count, "Calcul des temps de traitements pour chaque lot par centre."
        
    Next oFichierLotsCU
    
    Debug.Print "Calcul des cadences terminé"
End Sub

Private Function RécupérationCharge(centre As String, ProfilLOTPACK As String, QteLOTPACK As Integer) As Double
    Dim Temps As Object
    Dim Charge As Double
    
    ' Sélection du dictionnaire des temps en fonction du centre
    Select Case centre
        Case "DUBUS3"
            Set Temps = Temps_DUBUS3
        Case "DUBUS5"
            Set Temps = Temps_DUBUS5
        Case "DUBUS8"
            Set Temps = Temps_DUBUS8
        Case "SCHIRMER4"
            Set Temps = Temps_SCHIRMER4
        Case Else
            RécupérationCharge = -1 ' Centre non valide
            Exit Function
    End Select
    
    ' Vérification si le profil existe dans le dictionnaire des temps
    If Temps.Exists(ProfilLOTPACK) Then
        Charge = QteLOTPACK * Temps(ProfilLOTPACK) ' Calcul de la charge
    ' Si le profil complet n'existe pas, on tente avec les 5 premiers caractères
    ElseIf Temps.Exists(Left(ProfilLOTPACK, 5)) Then
        Charge = QteLOTPACK * Temps(Left(ProfilLOTPACK, 5))
    Else
        Charge = -1 ' Aucun temps trouvé pour ce profil
    End If
    
    ' Retourne la charge calculée
    RécupérationCharge = Charge
End Function

Private Function PénalitésCU(Profil As String, Charge As Double, centre As String, LotCU As String) As Double
    'Ce sub vise à Attribuer des pénalités à des profils qu'on souhaite autoriser sur les centres tout en tenant compte de contraintes
    Dim ProfilConcernés() As Variant
    Dim Pénalité As Double
    Dim ComparaisonProfils As Long
    
    PénalitésCU = 0
    
    Select Case centre
    Case "DUBUS5"
        ProfilConcernés = Array("PAM370", "PAM371", "PAM373", "PAM374", "PAM375", "PAM377") 'Profils dormants lourds
        Pénalité = 3
        
        For ComparaisonProfils = LBound(ProfilConcernés) To UBound(ProfilConcernés)
            If ProfilConcernés(ComparaisonProfils) = Profil Then
                PénalitésCU = Charge * Pénalité 'On ajoute 100% de charge pour matérialiser le passage de ces profils
                Exit For
            End If
        Next ComparaisonProfils

        If Dictionnaire_OtPE.Exists(LotCU) Then
            PénalitésCU = 1000000000#
        Else
            PénalitésCU = 0
        End If
        
    Case Else
    
    End Select

End Function

