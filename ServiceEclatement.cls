VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceEclatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' SERVICE D'ÉCLATEMENT DES LOTS
' ============================================================================
' Remplace : C_EclatementLots.EclatementLots()
'
' Responsabilité :
'   - Lire les fichiers d'usinage depuis le dossier S:\ALU\OPTIM DUBUS
'   - Parser chaque fichier (format DB/DP/ET/OP séparé par ";")
'   - Extraire les profils, barres, contraintes
'   - Créer les fichiers COMPIL_*.LOT dans %TEMP%
'   - Détecter les contraintes spéciales (RackOptim, OP668/669)
'   - Peupler les LotCU avec les infos détaillées
'
' Paramètres d'entrée :
'   - moteur.Lots (CollectionLots) peuplée par ServiceImportation
'   - datesProd (dict) : pour valider les lots BC
'
' Paramètres de sortie :
'   - Dictionnaires de suivi:
'     * profilsParLot : LotCU → Collection de profils
'     * barresParLot : LotCU → Collection de barres
'     * rackOptim : LotDestOPTIM → LotSourceOPTIM
'     * operationsInterdites : LotCU → true (OP668/669)
' ============================================================================

Private fileSystem As Object
Private cheminALU As String
Private cheminTemp As String
Private moteur As MoteurRepartition
Private logger As Object

' Index des fichiers ALU par lot (optimisation)
Private fichiersParLot As Object

' Dictionnaires de suivi (publics via getter)
Private profilsParLot As Object
Private barresParLot As Object
Private rackOptim As Object
Private operationsInterdites As Object
Private contenusLots As Object
Private fichierAvecExtension As Object

Public Property Get ProfilsParLot() As Object
    Set ProfilsParLot = profilsParLot
End Property

Public Property Get BarresParLot() As Object
    Set BarresParLot = barresParLot
End Property

Public Property Get RackOptim() As Object
    Set RackOptim = rackOptim
End Property

Public Property Get OperationsInterdites() As Object
    Set OperationsInterdites = operationsInterdites
End Property

Public Property Get ContenutsLots() As Object
    Set ContenutsLots = contenusLots
End Property

Public Property Get FichierAvecExtension() As Object
    Set FichierAvecExtension = fichierAvecExtension
End Property

' ============================================================================
' INITIALISATION
' ============================================================================

Public Sub Initialiser(moteurRef As MoteurRepartition, cheminALUFournisseur As String, cheminTempFournisseur As String, loggerRef As Object)
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set moteur = moteurRef
    Set logger = loggerRef
    cheminALU = cheminALUFournisseur
    cheminTemp = cheminTempFournisseur
    
    ' Initialiser les dictionnaires
    Set fichiersParLot = CreateObject("Scripting.Dictionary")
    Set profilsParLot = CreateObject("Scripting.Dictionary")
    Set barresParLot = CreateObject("Scripting.Dictionary")
    Set rackOptim = CreateObject("Scripting.Dictionary")
    Set operationsInterdites = CreateObject("Scripting.Dictionary")
    Set contenusLots = CreateObject("Scripting.Dictionary")
    Set fichierAvecExtension = CreateObject("Scripting.Dictionary")
    
    Log_Info "ServiceEclatement initialisé. Chemin ALU: " & cheminALU & ", Temp: " & cheminTemp
End Sub

' ============================================================================
' ENTRY POINT : Éclatement complet des lots
' ============================================================================

Public Sub EclatementLots(datesProd As Object)
    On Error GoTo ErrHandler
    
    Dim dossierALU As Object
    Dim fichierALU As Object
    Dim nombreFichiers As Long
    Dim indexFichier As Long
    Dim lotBC As Variant
    
    ' Vérifier les conditions préalables
    If moteur Is Nothing Then
        Err.Raise 2001, "ServiceEclatement.EclatementLots", "Service non initialisé."
    End If
    
    If Not fileSystem.FolderExists(cheminALU) Then
        Err.Raise 2002, "ServiceEclatement.EclatementLots", "Dossier ALU introuvable: " & cheminALU
    End If
    
    If Not fileSystem.FolderExists(cheminTemp) Then
        Err.Raise 2003, "ServiceEclatement.EclatementLots", "Dossier TEMP introuvable: " & cheminTemp
    End If
    
    ' Purger le dossier TEMP avant de commencer
    PurgerTemp
    
    Set dossierALU = fileSystem.GetFolder(cheminALU)
    nombreFichiers = dossierALU.Files.Count
    
    Log_Info "Chargement de " & nombreFichiers & " fichier(s) ALU..."
    
    ' ÉTAPE 1 : Pré-indexer les fichiers par lot BC
    indexFichier = 0
    For Each fichierALU In dossierALU.Files
        indexFichier = indexFichier + 1
        
        If moteur.UICallback Is Nothing Then
            Log_Debug "Pré-indexation " & indexFichier & "/" & nombreFichiers & ": " & fichierALU.Name
        Else
            moteur.UICallback.MajProgression indexFichier, nombreFichiers, "Pré-indexation fichiers ALU..."
        End If
        
        IndexerFichier fichierALU, datesProd
    Next fichierALU
    
    Log_Info "Pré-indexation terminée. " & fichiersParLot.Count & " lot(s) trouvé(s)."
    
    ' ÉTAPE 2 : Traiter chaque lot
    Dim iterLot As Long
    iterLot = 0
    
    For Each lotBC In datesProd.Keys
        iterLot = iterLot + 1
        
        If moteur.UICallback Is Nothing Then
            Log_Debug "Éclatement lot " & iterLot & "/" & datesProd.Count & ": " & lotBC
        Else
            moteur.UICallback.MajProgression iterLot, datesProd.Count, "Éclatement des lots: " & lotBC
        End If
        
        ' Si des fichiers correspondent à ce lot, les traiter
        If fichiersParLot.Exists(lotBC) Then
            TraiterLot lotBC, fichiersParLot(lotBC)
        Else
            Log_Debug "Aucun fichier trouvé pour le lot BC: " & lotBC
        End If
    Next lotBC
    
    Log_Info "Éclatement terminé. " & profilsParLot.Count & " lot(s) éclaté(s)."
    
    Exit Sub
ErrHandler:
    Log_Erreur "EclatementLots", Err.Description
    Err.Raise Err.Number
End Sub

' ============================================================================
' INDEXER LES FICHIERS PAR LOT BC
' ============================================================================

Private Sub IndexerFichier(fichierALU As Object, datesProd As Object)
    On Error Resume Next ' Gestion souple pour les fichiers non reconnus
    
    Dim nomFichier As String
    Dim lotBC As Variant
    Dim collFichiers As Collection
    
    nomFichier = fichierALU.Name
    
    ' Tester si ce fichier appartient à un des lots BC
    For Each lotBC In datesProd.Keys
        If nomFichier Like lotBC & "*" Then
            ' Créer la collection si elle n'existe pas
            If Not fichiersParLot.Exists(lotBC) Then
                Set collFichiers = New Collection
                fichiersParLot.Add lotBC, collFichiers
            End If
            
            ' Ajouter le fichier à la collection du lot
            fichiersParLot(lotBC).Add fichierALU
            Exit Sub ' Un fichier ne peut appartenir qu'à un seul lot
        End If
    Next lotBC
    
    Log_Debug "Fichier non indexé (probablement ancien ou non reconnu): " & nomFichier
    On Error GoTo 0
End Sub

' ============================================================================
' TRAITER UN LOT (Parser tous ses fichiers d'usinage)
' ============================================================================

Private Sub TraiterLot(lotBC As String, collFichiersLot As Collection)
    On Error GoTo ErrHandler
    
    Dim fichierALU As Object
    Dim lotCU As String
    Dim contenuLot() As String
    Dim fichierLOT As Object
    Dim contenuBrut As String
    Dim i As Long
    Dim ligneLot As String
    Dim champsLot() As String
    Dim profilLot As String
    Dim clefBarre As String
    Dim typeLot As String
    Dim dictProfils As Object
    Dim dictBarres As Object
    
    Set dictProfils = CreateObject("Scripting.Dictionary")
    Set dictBarres = CreateObject("Scripting.Dictionary")
    
    ' Parser chaque fichier du lot
    For Each fichierALU In collFichiersLot
        
        ' Lire le fichier
        Set fichierLOT = fileSystem.OpenTextFile(fichierALU.Path, 1)
        contenuBrut = fichierLOT.ReadAll
        fichierLOT.Close
        
        ' Extraire le nom du lot CU (sans extension)
        lotCU = Left(fichierALU.Name, InStrRev(fichierALU.Name, ".") - 1)
        fichierAvecExtension(lotCU) = fichierALU.Name
        
        ' Déterminer le type du lot
        typeLot = DeterminerTypeDeLot(lotCU)
        
        ' Découper en lignes
        contenuLot = Split(contenuBrut, vbNewLine)
        
        ' Parser chaque ligne
        For i = LBound(contenuLot) To UBound(contenuLot)
            ligneLot = Trim(contenuLot(i))
            
            If Len(ligneLot) > 0 Then
                champsLot = Split(ligneLot, ";")
                
                If UBound(champsLot) >= 2 Then
                    Select Case champsLot(0)
                        
                        Case "DB"  ' Début de barre
                            profilLot = Replace(champsLot(1), """", "")
                            clefBarre = IIf(Len(sequencesBC) > 0, sequencesBC(lotBC), "SEQ???") & "-" & profilLot & "$" & Replace(champsLot(2), """", "") & "$" & Replace(champsLot(3), """", "") & "-" & lotCU
                            
                            ' Compter les barres
                            If Not dictBarres.Exists(clefBarre) Then
                                dictBarres.Add clefBarre, 1
                            Else
                                dictBarres(clefBarre) = dictBarres(clefBarre) + 1
                            End If
                        
                        Case "DP"  ' Début de pièce
                            ' Compter les profils
                            If Not dictProfils.Exists(profilLot) Then
                                dictProfils.Add profilLot, 1
                            Else
                                dictProfils(profilLot) = dictProfils(profilLot) + 1
                            End If
                        
                        Case "ET"  ' Étiquette - détection RackOptim
                            If UBound(champsLot) >= 10 Then
                                ' Si le lot de l'étiquette != lot du fichier → RackOptim détecté
                                If Not lotCU Like champsLot(10) & "*" Then
                                    rackOptim.Add champsLot(10) & typeLot, lotCU
                                    Log_Debug "RackOptim détecté: " & champsLot(10) & " source de " & lotCU
                                End If
                            End If
                        
                        Case "OP"  ' Opération - détection OP668/669
                            If UBound(champsLot) >= 1 Then
                                If champsLot(1) = "668" Or champsLot(1) = "669" Then
                                    operationsInterdites.Add lotCU, True
                                    Log_Debug "Opération interdite (OP" & champsLot(1) & ") détectée sur: " & lotCU
                                End If
                            End If
                    
                    End Select
                End If
            End If
        Next i
        
        ' Écrire le fichier COMPIL_*.LOT
        EcrireCompilLot lotCU, dictProfils, typeLot
        
        ' Stocker les données
        If profilsParLot.Count = 0 Or Not profilsParLot.Exists(lotCU) Then
            profilsParLot.Add lotCU, dictProfils
        End If
        
    Next fichierALU
    
    ' Stocker les barres
    Dim clef As Variant
    For Each clef In dictBarres.Keys
        If Not barresParLot.Exists(clef) Then
            barresParLot.Add clef, dictBarres(clef)
        End If
    Next clef
    
    Exit Sub
ErrHandler:
    Log_Erreur "TraiterLot", Err.Description
    ' Continuer malgré l'erreur pour ce lot
    On Error Resume Next
End Sub

' ============================================================================
' ÉCRIRE LE FICHIER COMPIL_*.LOT
' ============================================================================

Private Sub EcrireCompilLot(lotCU As String, dictProfils As Object, typeLot As String)
    On Error GoTo ErrHandler
    
    Dim cheminCompil As String
    Dim fichierCompil As Object
    Dim profil As Variant
    
    cheminCompil = cheminTemp & "COMPIL_" & lotCU & ".LOT"
    
    ' Créer le fichier
    Set fichierCompil = fileSystem.CreateTextFile(cheminCompil, True)
    
    ' Écrire chaque profil et sa quantité
    For Each profil In dictProfils.Keys
        fichierCompil.WriteLine profil & ";" & dictProfils(profil)
        
        ' Aussi stocker dans le dictionnaire de suivi
        contenusLots.Add lotCU & ";" & profil, dictProfils(profil)
    Next profil
    
    fichierCompil.Close
    
    Log_Debug "Fichier COMPIL créé: " & cheminCompil
    
    Exit Sub
ErrHandler:
    Log_Erreur "EcrireCompilLot", Err.Description
    On Error Resume Next
End Sub

' ============================================================================
' PURGER LE DOSSIER TEMP
' ============================================================================

Private Sub PurgerTemp()
    On Error GoTo ErrHandler
    
    Dim dossierTemp As Object
    Dim fichier As Object
    Dim nombreFichiers As Long
    
    If Not fileSystem.FolderExists(cheminTemp) Then
        ' Créer le dossier s'il n'existe pas
        fileSystem.CreateFolder cheminTemp
        Exit Sub
    End If
    
    Set dossierTemp = fileSystem.GetFolder(cheminTemp)
    
    ' Supprimer tous les fichiers COMPIL_*.LOT
    For Each fichier In dossierTemp.Files
        If fichier.Name Like "COMPIL_*.LOT" Then
            fichier.Delete
        End If
    Next fichier
    
    Log_Debug "Dossier TEMP purgé"
    
    Exit Sub
ErrHandler:
    Log_Erreur "PurgerTemp", Err.Description
    ' Non bloquant
    On Error Resume Next
End Sub

' ============================================================================
' DÉTERMINER LE TYPE D'UN LOT
' ============================================================================

Private Function DeterminerTypeDeLot(nomDuLot As String) As String
    Dim typesLots As Variant
    Dim i As Long
    
    typesLots = Array("AUTO", "COL", "COLIS", "COP", "COPCOUL", "DOR", "FBC030", "OUV", _
                      "PAM208", "PAM207", "PAM305", "PAM306", "PAM322", "PARCLOSE", "Scie", "VR")
    
    For i = LBound(typesLots) To UBound(typesLots)
        If Right(nomDuLot, Len(typesLots(i))) = typesLots(i) Then
            DeterminerTypeDeLot = typesLots(i)
            Exit Function
        End If
    Next i
    
    DeterminerTypeDeLot = ""
End Function

' ============================================================================
' UTILITAIRES DE LOGGING
' ============================================================================

Private Sub Log_Info(message As String)
    If logger Is Nothing Then
        Debug.Print "[ServiceEclatement] INFO: " & message
    Else
        On Error Resume Next
        logger.LogInfo message
        On Error GoTo 0
    End If
End Sub

Private Sub Log_Debug(message As String)
    If logger Is Nothing Then
        Debug.Print "[ServiceEclatement] DEBUG: " & message
    Else
        On Error Resume Next
        logger.LogDebug message
        On Error GoTo 0
    End If
End Sub

Private Sub Log_Erreur(nomSub As String, message As String)
    Dim msg As String
    msg = "[ServiceEclatement::" & nomSub & "] ERREUR: " & message
    
    If logger Is Nothing Then
        Debug.Print msg
    Else
        On Error Resume Next
        logger.LogErreur msg
        On Error GoTo 0
    End If
End Sub

' Variable globale helper (sera définie dans MoteurOrchestration)
Private sequencesBC As Object
