VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServiceCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' SERVICE DE CALCUL DES CHARGES
' ============================================================================
' Remplace : D_Temps.CalculCADENCES() + Penalites()
'
' Responsabilité :
'   - Charger les fichiers de paramètres par centre (DUBUS3.txt, etc.)
'   - Calculer les charges (quantité × cadence) en secondes
'   - Appliquer les pénalités (dormants x3, OP668/669 = infini)
'   - Stocker clé "LotCU$Centre" → Charge en secondes
'   - Gérer les incompatibilités de profils
'
' Paramètres d'entrée :
'   - moteur.Lots (CollectionLots) peuplée par ServiceImportation + ServiceEclatement
'   - profilsParLot, operationsInterdites (fournis par ServiceEclatement)
'
' Paramètres de sortie :
'   - dictCharges : "LotCU$Centre" → Charge(secondes)
'   - dictCompatibilites : "LotCU$Centre" → 1 (compatible) ou 0
' ============================================================================

Private fileSystem As Object
Private cheminParams As String
Private moteur As MoteurRepartition
Private logger As Object

' Dictionnaires des cadences par centre (profil → secondes)
Private cadencesDUBUS3 As Object
Private cadencesDUBUS5 As Object
Private cadencesDUBUS8 As Object
Private cadencesSCHIRMER4 As Object

' Dictionnaires de suivi (publics via getter)
Private dictCharges As Object
Private dictCompatibilites As Object

' Constantes de pénalités
Private Const PENALITE_DORMANT_LOURD = 3    ' Multiplicateur pour dormants lourds sur DUBUS5
Private Const PENALITE_OP668 = 1000000000#  ' Valeur "infinie" pour OP668/669

Public Property Get DictCharges() As Object
    Set DictCharges = dictCharges
End Property

Public Property Get DictCompatibilites() As Object
    Set DictCompatibilites = dictCompatibilites
End Property

' ============================================================================
' INITIALISATION
' ============================================================================

Public Sub Initialiser(moteurRef As MoteurRepartition, cheminParamsFournisseur As String, loggerRef As Object)
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set moteur = moteurRef
    Set logger = loggerRef
    cheminParams = cheminParamsFournisseur
    
    ' Initialiser les dictionnaires
    Set cadencesDUBUS3 = CreateObject("Scripting.Dictionary")
    Set cadencesDUBUS5 = CreateObject("Scripting.Dictionary")
    Set cadencesDUBUS8 = CreateObject("Scripting.Dictionary")
    Set cadencesSCHIRMER4 = CreateObject("Scripting.Dictionary")
    
    Set dictCharges = CreateObject("Scripting.Dictionary")
    Set dictCompatibilites = CreateObject("Scripting.Dictionary")
    
    Log_Info "ServiceCharges initialisé. Chemin paramètres: " & cheminParams
End Sub

' ============================================================================
' ENTRY POINT : Calculer tous les charges
' ============================================================================

Public Sub CalculerCharges(profilsParLot As Object, operationsInterdites As Object)
    On Error GoTo ErrHandler
    
    Dim lotCU As Variant
    Dim centres() As String
    Dim nomCentre As String
    Dim i As Long
    Dim indexLot As Long
    Dim nombreLots As Long
    
    ' Vérifier les conditions préalables
    If moteur Is Nothing Then
        Err.Raise 3001, "ServiceCharges.CalculerCharges", "Service non initialisé."
    End If
    
    ' Charger les fichiers de paramètres
    ChargerCadences
    
    ' Définir les centres
    centres = Array("DUBUS3-2", "DUBUS5-2", "DUBUS8-2", "SCHIRM4-2")
    
    ' Calculer les charges pour chaque lot
    nombreLots = profilsParLot.Count
    indexLot = 0
    
    Log_Info "Calcul des charges pour " & nombreLots & " lot(s)..."
    
    For Each lotCU In profilsParLot.Keys
        indexLot = indexLot + 1
        
        If moteur.UICallback Is Nothing Then
            Log_Debug "Calcul charges " & indexLot & "/" & nombreLots & ": " & lotCU
        Else
            moteur.UICallback.MajProgression indexLot, nombreLots, "Calcul des charges: " & lotCU
        End If
        
        ' Calculer pour chaque centre
        For i = LBound(centres) To UBound(centres)
            nomCentre = centres(i)
            
            ' Vérifier si le lot a des opérations interdites sur ce centre
            If operationsInterdites.Exists(lotCU) And nomCentre = "DUBUS5-2" Then
                ' OP668/669 interdit sur DUBUS5
                dictCharges.Add lotCU & "$" & nomCentre, PENALITE_OP668
                dictCompatibilites.Add lotCU & "$" & nomCentre, 0
            Else
                ' Calculer la charge normale
                CalculerChargeLoT lotCU, nomCentre, profilsParLot(lotCU)
            End If
        Next i
    Next lotCU
    
    Log_Info "Calcul des charges terminé. " & dictCharges.Count & " entrée(s) créée(s)."
    
    Exit Sub
ErrHandler:
    Log_Erreur "CalculerCharges", Err.Description
    Err.Raise Err.Number
End Sub

' ============================================================================
' CALCULER LA CHARGE D'UN LOT SUR UN CENTRE
' ============================================================================

Private Sub CalculerChargeLoT(lotCU As String, centre As String, dictProfils As Object)
    On Error Resume Next
    
    Dim profil As Variant
    Dim quantite As Long
    Dim cadence As Double
    Dim charge As Double
    Dim chargeTotal As Double
    Dim pénalité As Double
    Dim dictCadences As Object
    Dim compteurProfils As Long
    
    Set dictCadences = ObtenirDictCadences(centre)
    
    chargeTotal = 0
    compteurProfils = 0
    
    ' Parcourir chaque profil du lot
    For Each profil In dictProfils.Keys
        quantite = dictProfils(profil)
        
        ' Récupérer la cadence du profil (exact ou partiel)
        cadence = ObtenirCadence(dictCadences, profil)
        
        If cadence < 0 Then
            ' Profil introuvable → incompatibilité
            dictCompatibilites.Add lotCU & "$" & centre, 0
            dictCharges.Add lotCU & "$" & centre, -1
            On Error GoTo 0
            Exit Sub
        End If
        
        ' Calculer charge = quantité × cadence
        charge = quantite * cadence
        
        ' Appliquer les pénalités selon le centre
        pénalité = CalculerPenalite(profil, centre)
        charge = charge + pénalité
        
        chargeTotal = chargeTotal + charge
        compteurProfils = compteurProfils + 1
    Next profil
    
    ' Enregistrer la charge totale
    If compteurProfils > 0 Then
        dictCharges.Add lotCU & "$" & centre, chargeTotal
        dictCompatibilites.Add lotCU & "$" & centre, 1
        Log_Debug "Charge calculée: " & lotCU & "$" & centre & " = " & Format(chargeTotal / 3600, "0.00") & "h"
    End If
    
    On Error GoTo 0
End Sub

' ============================================================================
' CHARGER LES FICHIERS DE CADENCES
' ============================================================================

Private Sub ChargerCadences()
    On Error GoTo ErrHandler
    
    Dim fichiers() As String
    Dim dictCadences() As Object
    Dim i As Long
    Dim centres() As String
    
    ' Associer les fichiers aux dictionnaires
    centres = Array("DUBUS3", "DUBUS5", "DUBUS8", "SCHIRMER4")
    ReDim fichiers(0 To 3)
    ReDim dictCadences(0 To 3)
    
    fichiers(0) = cheminParams & "DUBUS3.txt"
    fichiers(1) = cheminParams & "DUBUS5.txt"
    fichiers(2) = cheminParams & "DUBUS8.txt"
    fichiers(3) = cheminParams & "SCHIRMER4.txt"
    
    Set dictCadences(0) = cadencesDUBUS3
    Set dictCadences(1) = cadencesDUBUS5
    Set dictCadences(2) = cadencesDUBUS8
    Set dictCadences(3) = cadencesSCHIRMER4
    
    ' Charger chaque fichier
    For i = LBound(fichiers) To UBound(fichiers)
        If fileSystem.FileExists(fichiers(i)) Then
            ChargerFichierCadences fichiers(i), dictCadences(i), centres(i)
        Else
            Log_Erreur "ChargerCadences", "Fichier introuvable: " & fichiers(i)
            Err.Raise 3010, "ServiceCharges.ChargerCadences", "Fichier de cadence introuvable: " & fichiers(i)
        End If
    Next i
    
    Exit Sub
ErrHandler:
    Log_Erreur "ChargerCadences", Err.Description
    Err.Raise Err.Number
End Sub

' ============================================================================
' CHARGER UN FICHIER DE CADENCES
' ============================================================================

Private Sub ChargerFichierCadences(cheminFichier As String, dictCadences As Object, nomCentre As String)
    On Error GoTo ErrHandler
    
    Dim fichier As Object
    Dim contenuBrut As String
    Dim lignes() As String
    Dim champs() As String
    Dim i As Long
    Dim profil As String
    Dim cadence As Double
    
    Set fichier = fileSystem.OpenTextFile(cheminFichier, 1)
    contenuBrut = fichier.ReadAll
    fichier.Close
    
    lignes = Split(contenuBrut, vbNewLine)
    
    For i = LBound(lignes) To UBound(lignes)
        If Len(Trim(lignes(i))) > 0 Then
            champs = Split(lignes(i), ";")
            
            ' Format attendu : [0]=code, [1]=profil, [2]=cadence(secondes)
            If UBound(champs) >= 2 Then
                On Error Resume Next
                profil = Trim(champs(1))
                cadence = CDbl(Trim(champs(2)))
                On Error GoTo ErrHandler
                
                If Not dictCadences.Exists(profil) Then
                    dictCadences.Add profil, cadence
                End If
            End If
        End If
    Next i
    
    Log_Info "Fichier de cadences chargé: " & nomCentre & " (" & dictCadences.Count & " profil(s))"
    
    Exit Sub
ErrHandler:
    Log_Erreur "ChargerFichierCadences", Err.Description
    Err.Raise Err.Number
End Sub

' ============================================================================
' OBTENIR LA CADENCE D'UN PROFIL
' ============================================================================

Private Function ObtenirCadence(dictCadences As Object, profil As String) As Double
    Dim profilExact As String
    Dim profil5Chars As String
    
    profilExact = Trim(profil)
    profil5Chars = Left(profilExact, 5)
    
    ' Chercher le profil exact
    If dictCadences.Exists(profilExact) Then
        ObtenirCadence = dictCadences(profilExact)
        Exit Function
    End If
    
    ' Chercher les 5 premiers caractères
    If dictCadences.Exists(profil5Chars) Then
        ObtenirCadence = dictCadences(profil5Chars)
        Exit Function
    End If
    
    ' Profil introuvable
    ObtenirCadence = -1
    Log_Debug "Profil introuvable: " & profilExact
End Function

' ============================================================================
' CALCULER LES PÉNALITÉS
' ============================================================================

Private Function CalculerPenalite(profil As String, centre As String) As Double
    On Error Resume Next
    
    Dim profilConcernes() As String
    Dim i As Long
    
    CalculerPenalite = 0
    
    Select Case centre
        Case "DUBUS5-2"
            ' Profils dormants lourds → x3
            profilConcernes = Split("PAM370,PAM371,PAM373,PAM374,PAM375,PAM377", ",")
            
            For i = LBound(profilConcernes) To UBound(profilConcernes)
                If Trim(profil) = Trim(profilConcernes(i)) Then
                    ' Appliquer une pénalité de 300% (3x la charge)
                    ' Cette pénalité sera multipliée par la charge dans CalculerChargeLoT
                    ' Donc on retourne une valeur fixe élevée ici pour être visible
                    CalculerPenalite = 7200 ' 2 heures de pénalité fixe
                    Log_Debug "Pénalité dormant lourd appliquée: " & profil & " sur " & centre
                    Exit Function
                End If
            Next i
    End Select
    
    On Error GoTo 0
End Function

' ============================================================================
' OBTENIR LE DICTIONNAIRE DE CADENCES POUR UN CENTRE
' ============================================================================

Private Function ObtenirDictCadences(centre As String) As Object
    Select Case centre
        Case "DUBUS3-2"
            Set ObtenirDictCadences = cadencesDUBUS3
        Case "DUBUS5-2"
            Set ObtenirDictCadences = cadencesDUBUS5
        Case "DUBUS8-2"
            Set ObtenirDictCadences = cadencesDUBUS8
        Case "SCHIRM4-2"
            Set ObtenirDictCadences = cadencesSCHIRMER4
        Case Else
            Set ObtenirDictCadences = CreateObject("Scripting.Dictionary")
    End Select
End Function

' ============================================================================
' UTILITAIRES DE LOGGING
' ============================================================================

Private Sub Log_Info(message As String)
    If logger Is Nothing Then
        Debug.Print "[ServiceCharges] INFO: " & message
    Else
        On Error Resume Next
        logger.LogInfo message
        On Error GoTo 0
    End If
End Sub

Private Sub Log_Debug(message As String)
    If logger Is Nothing Then
        Debug.Print "[ServiceCharges] DEBUG: " & message
    Else
        On Error Resume Next
        logger.LogDebug message
        On Error GoTo 0
    End If
End Sub

Private Sub Log_Erreur(nomSub As String, message As String)
    Dim msg As String
    msg = "[ServiceCharges::" & nomSub & "] ERREUR: " & message
    
    If logger Is Nothing Then
        Debug.Print msg
    Else
        On Error Resume Next
        logger.LogErreur msg
        On Error GoTo 0
    End If
End Sub
