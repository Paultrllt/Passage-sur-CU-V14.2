VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceImportation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' SERVICE D'IMPORTATION BC
' ============================================================================
' Remplace : B_AnalyseBC.RecuperationBesoinBC() + B_AnalyseBC.LectureExportBC()
' 
' Responsabilité :
'   - Lire les fichiers d'export BC depuis le dossier 01_ExportBC
'   - Parser les lignes au format délimité par ";"
'   - Créer des objets LotCU et les ajouter à la collection
'   - Gérer les dates de production et les jours de fabrication
'   - Calculer les informations de synthèse (séquence, ouvrants, dormants)
'
' Paramètres de sortie :
'   - moteur.Lots (CollectionLots) peuplée avec les LotCU créés
'   - Dictionnaires de suivi :
'     * datesProd (dict) : LotBC → DateProduction
'     * sequencesBC (dict) : LotBC → Séquence MES
'     * syntheseProd (dict) : LotBC → "Seq;NbOuv;NbDorm"
'   - Jours de production (1-5) détectés
' ============================================================================

Private fileSystem As Object
Private cheminBC As String
Private moteur As MoteurRepartition
Private logger As Object

' Dictionnaires de suivi (publics via getter)
Private datesProd As Object
Private sequencesBC As Object
Private syntheseProd As Object
Private lotsDateUniques As Object

' Variables de résultat
Private nombreJours As Long
Private joursProduction() As Date
Private semaineExp As String
Private sfab As Integer
Private annee As Integer

Public Property Get DatesProd() As Object
    Set DatesProd = datesProd
End Property

Public Property Get SequencesBC() As Object
    Set SequencesBC = sequencesBC
End Property

Public Property Get SyntheseProd() As Object
    Set SyntheseProd = syntheseProd
End Property

Public Property Get NombreJours() As Long
    NombreJours = nombreJours
End Property

Public Property Get JoursProduction() As Date()
    JoursProduction = joursProduction
End Property

Public Property Get SemaineExp() As String
    SemaineExp = semaineExp
End Property

Public Property Get SFAB() As Integer
    SFAB = sfab
End Property

Public Property Get Annee() As Integer
    Annee = annee
End Property

' ============================================================================
' INITIALISATION
' ============================================================================

Public Sub Initialiser(moteurRef As MoteurRepartition, cheminBCFournisseur As String, loggerRef As Object)
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set moteur = moteurRef
    Set logger = loggerRef
    cheminBC = cheminBCFournisseur
    
    ' Initialiser les dictionnaires
    Set datesProd = CreateObject("Scripting.Dictionary")
    Set sequencesBC = CreateObject("Scripting.Dictionary")
    Set syntheseProd = CreateObject("Scripting.Dictionary")
    Set lotsDateUniques = CreateObject("Scripting.Dictionary")
    
    ReDim joursProduction(1 To 5)
    
    Log_Info "ServiceImportation initialisé. Chemin BC: " & cheminBC
End Sub

' ============================================================================
' ENTRY POINT : Charger tous les fichiers BC
' ============================================================================

Public Sub ChargerFichiersBC()
    On Error GoTo ErrHandler
    
    Dim dossierBC As Object
    Dim fichierBC As Object
    Dim nombreFichiers As Long
    Dim indexFichier As Long
    
    ' Vérifier que le service est initialisé
    If moteur Is Nothing Then
        Err.Raise 1001, "ServiceImportation.ChargerFichiersBC", "Service non initialisé. Appelez Initialiser() d'abord."
    End If
    
    ' Accéder au dossier BC
    If Not fileSystem.FolderExists(cheminBC) Then
        Err.Raise 1002, "ServiceImportation.ChargerFichiersBC", "Dossier BC introuvable: " & cheminBC
    End If
    
    Set dossierBC = fileSystem.GetFolder(cheminBC)
    nombreFichiers = dossierBC.Files.Count
    
    If nombreFichiers = 0 Then
        Err.Raise 1003, "ServiceImportation.ChargerFichiersBC", "Aucun fichier BC trouvé dans: " & cheminBC
    End If
    
    Log_Info "Chargement de " & nombreFichiers & " fichier(s) BC..."
    
    ' Traiter chaque fichier BC (généralement 1, mais peut être plusieurs)
    indexFichier = 0
    For Each fichierBC In dossierBC.Files
        indexFichier = indexFichier + 1
        
        ' Afficher progression
        If moteur.UICallback Is Nothing Then
            ' Pas d'UI, on log simplement
            Log_Info "Traitement fichier " & indexFichier & "/" & nombreFichiers & ": " & fichierBC.Name
        Else
            ' Appeler le callback UI
            moteur.UICallback.MajProgression indexFichier, nombreFichiers, "Import BC: " & fichierBC.Name
        End If
        
        ' Parser le fichier BC
        ParserFichierBC fichierBC.Path
    Next fichierBC
    
    ' Calculer les jours de production à partir des dates uniques
    CalculerJoursProduction
    
    Log_Info "Importation BC terminée. " & moteur.Lots.Count & " lot(s) créé(s). " & nombreJours & " jour(s) détecté(s)."
    
    Exit Sub
ErrHandler:
    Log_Erreur "ChargerFichiersBC", Err.Description
    Err.Raise Err.Number
End Sub

' ============================================================================
' PARSER UN FICHIER BC
' ============================================================================

Private Sub ParserFichierBC(cheminFichier As String)
    On Error GoTo ErrHandler
    
    Dim fichier As Object
    Dim contenuBrut As String
    Dim lignes() As String
    Dim i As Long
    Dim champs() As String
    Dim lotBC As String
    Dim dateProduction As Date
    Dim sequence As String
    Dim nbOuvrants As String
    Dim nbDormants As String
    Dim nombreLignes As Long
    
    ' Ouvrir et lire le fichier
    Set fichier = fileSystem.OpenTextFile(cheminFichier, 1)
    contenuBrut = fichier.ReadAll
    fichier.Close
    
    ' Découper en lignes
    lignes = Split(contenuBrut, vbNewLine)
    nombreLignes = UBound(lignes) - LBound(lignes) + 1
    
    Log_Debug "Fichier BC contient " & nombreLignes & " ligne(s)"
    
    ' Parser chaque ligne
    For i = LBound(lignes) To UBound(lignes)
        If Len(Trim(lignes(i))) > 0 Then
            champs = Split(lignes(i), ";")
            
            ' Format attendu : [0]=DateProd, [1]=???, [2]=Séquence, [3]=NbOuv, [4]=NbDorm, [6]=SemExp, [7]=CodeLot
            If UBound(champs) >= 7 And champs(7) <> "Code lot" Then
                
                On Error Resume Next ' Gestion souple des erreurs de conversion
                
                lotBC = champs(7)
                dateProduction = CDate(champs(1))
                sequence = champs(2)
                nbOuvrants = champs(3)
                nbDormants = champs(4)
                semaineExp = champs(6)
                
                ' Enregistrer les infos BC
                datesProd(lotBC) = dateProduction
                sequencesBC(lotBC) = sequence
                syntheseProd(lotBC) = sequence & ";" & nbOuvrants & ";" & nbDormants
                
                ' Enregistrer la date unique
                If Not lotsDateUniques.Exists(CStr(dateProduction)) Then
                    lotsDateUniques.Add CStr(dateProduction), dateProduction
                End If
                
                ' Récupérer SFAB et Année à partir de la première date valide
                If sfab = 0 Then
                    sfab = DatePart("ww", dateProduction, vbMonday, vbFirstFourDays)
                    annee = Year(dateProduction)
                    Log_Debug "SFAB détecté: " & sfab & ", Année: " & annee
                End If
                
                On Error GoTo ErrHandler
                
            End If
        End If
    Next i
    
    Exit Sub
ErrHandler:
    Log_Erreur "ParserFichierBC", Err.Description
    Err.Raise Err.Number
End Sub

' ============================================================================
' CALCULER LES JOURS DE PRODUCTION
' ============================================================================

Private Sub CalculerJoursProduction()
    On Error GoTo ErrHandler
    
    Dim dates() As Date
    Dim i As Long
    Dim index As Long
    Dim datesTriees() As Date
    
    ' Créer un tableau à partir des dates uniques
    ReDim dates(1 To lotsDateUniques.Count)
    index = 1
    
    For Each i In lotsDateUniques.Keys
        dates(index) = lotsDateUniques(i)
        index = index + 1
    Next i
    
    ' Trier les dates (QuickSort)
    If lotsDateUniques.Count > 0 Then
        TriRapideDate dates, LBound(dates), UBound(dates)
    End If
    
    ' Remplir le tableau joursProduction (max 5 jours)
    nombreJours = Application.WorksheetFunction.Min(5, lotsDateUniques.Count)
    
    For i = 1 To nombreJours
        joursProduction(i) = dates(i)
    Next i
    
    Log_Info nombreJours & " jour(s) de production détecté(s)"
    
    Exit Sub
ErrHandler:
    Log_Erreur "CalculerJoursProduction", Err.Description
    Err.Raise Err.Number
End Sub

' ============================================================================
' TRI RAPIDE POUR LES DATES
' ============================================================================

Private Sub TriRapideDate(arr() As Date, low As Long, high As Long)
    Dim pivot As Date, i As Long, j As Long, temp As Date
    
    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low
        j = high
        
        Do While i <= j
            Do While arr(i) < pivot
                i = i + 1
            Loop
            Do While arr(j) > pivot
                j = j - 1
            Loop
            
            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        
        TriRapideDate arr, low, j
        TriRapideDate arr, i, high
    End If
End Sub

' ============================================================================
' UTILITAIRES DE LOGGING
' ============================================================================

Private Sub Log_Info(message As String)
    If logger Is Nothing Then
        Debug.Print "[ServiceImportation] INFO: " & message
    Else
        On Error Resume Next
        logger.LogInfo message
        On Error GoTo 0
    End If
End Sub

Private Sub Log_Debug(message As String)
    If logger Is Nothing Then
        Debug.Print "[ServiceImportation] DEBUG: " & message
    Else
        On Error Resume Next
        logger.LogDebug message
        On Error GoTo 0
    End If
End Sub

Private Sub Log_Erreur(nomSub As String, message As String)
    Dim msg As String
    msg = "[ServiceImportation::" & nomSub & "] ERREUR: " & message
    
    If logger Is Nothing Then
        Debug.Print msg
    Else
        On Error Resume Next
        logger.LogErreur msg
        On Error GoTo 0
    End If
End Sub
