Attribute VB_Name = "C_EclatementLots"
Public Sub EclatementLots()
    ' Déclaration et initialisation des dictionnaires
    Dim Dictionnaire_PiecesCU As Object, Dictionnaire_Fichiers As Object
    Dim Dictionnaire_FichiersParLot As Object  ' Nouveau dictionnaire pour indexer les fichiers par lot
    
    Set Dictionnaire_PiecesCU = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_BarresCU = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_Fichiers = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_ContenuLot = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_FichiersParLot = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_OtPE = CreateObject("Scripting.Dictionary")
    
    Dim oFichierALU As Object
    Dim LotBC As Variant, LotCU As String
    Dim Regroupement As Boolean
    Dim TypeDuLot As String
    Dim CheminCHARGELOT As String
    Dim FichierCHARGELOT As Object
    Dim Profil As Variant
    
    ' Variables de suivi de la progression
    Dim Niter As Long, totalIter As Long
    EtapeMessage = "Recherche des fichiers d'usinage."
    Niter = 0
    
    ' --------------------------------------------------
    ' Chargement des fichiers depuis le dossier source
    ' --------------------------------------------------
    Dim oDossierALU As Object
    Set oDossierALU = Manipulationfichiers.GetFolder("S:\ALU\OPTIM DUBUS")
    totalIter = oDossierALU.Files.count
    For Each oFichierALU In oDossierALU.Files
        Niter = Niter + 1
        If Niter Mod 50 = 0 Then Configurateur.Maj_Progression
        Dictionnaire_Fichiers.Add oFichierALU.Name, oFichierALU
    Next oFichierALU
    Debug.Print "Chargement des fichiers terminé. Total fichiers: " & Dictionnaire_Fichiers.count
    
    ' --------------------------------------------------
    ' Pré-indexation des fichiers par lot
    ' --------------------------------------------------
    ' Pour chaque lot existant, on crée une collection vide dans Dictionnaire_FichiersParLot
    Dim nomLot As Variant
    Dim colFichiers As Collection
    For Each nomLot In Dictionnaire_LotsBC.Keys
        Set colFichiers = New Collection
        Dictionnaire_FichiersParLot.Add nomLot, colFichiers
    Next nomLot
    
    ' Parcours de tous les fichiers pour les associer au bon lot
    Dim nomFichier As Variant
    Dim fichierActuel As Object
    Dim trouveLot As Boolean
    For Each nomFichier In Dictionnaire_Fichiers.Keys
        Set fichierActuel = Dictionnaire_Fichiers(nomFichier)
        trouveLot = False
        ' On teste si le nom du fichier commence par le numéro du lot
        For Each nomLot In Dictionnaire_LotsBC.Keys
            If nomFichier Like nomLot & "*" Then
                Dictionnaire_FichiersParLot(nomLot).Add fichierActuel
                trouveLot = True
                Exit For  ' Un fichier ne peut appartenir qu'à un seul lot
            End If
        Next nomLot
        If Not trouveLot Then
            Debug.Print "Fichier non associé à aucun lot: " & nomFichier
        End If
    Next nomFichier
    
    ' --------------------------------------------------
    ' Traitement des lots
    ' --------------------------------------------------
    For Each LotBC In Dictionnaire_LotsBC.Keys
        Niter = 0
        Niter = Niter + 1
        If Niter Mod 10 = 0 Then Configurateur.Maj_Progression Niter, Dictionnaire_LotsBC.count, "Eclatement des lots."
        
        ' Récupérer la collection de fichiers correspondant au lot courant
        Dim colFichiersLot As Collection
        Set colFichiersLot = Dictionnaire_FichiersParLot(LotBC)
        
        ' Si aucun fichier ne correspond, on log et on passe au lot suivant
        If colFichiersLot.count = 0 Then
            Debug.Print "Aucune correspondance trouvée pour le lot: " & LotBC
            GoTo LotSuivant
        End If
        
        Dim oFichier As Object
        For Each oFichier In colFichiersLot
        
        TypeDuLot = DeterminerTypeDeLot(Left(oFichier.Name, Len(oFichier.Name) - 4))
            
            ' Lecture du contenu du fichier
            Dim ContenuLot() As String
            Dim ContenBRUT As String
            Dim FichierLOT As Object
            Set FichierLOT = Manipulationfichiers.OpenTextFile(oFichier, 1)
            ContenuBRUT = FichierLOT.readall
            ContenuLot = Split(ContenuBRUT, vbNewLine)
            FichierLOT.Close
            
            ' Extraction du nom du lot CU (suppression de l'extension)
            LotCU = Left(oFichier.Name, InStrRev(oFichier.Name, ".") - 1)
            Dictionnaire_FichierAvecExtension(LotCU) = oFichier.Name
            
            Dim i As Long, LigneLot As String, ChampsLot() As String, ProfilLot As String, ClefBarre As String
            For i = LBound(ContenuLot) To UBound(ContenuLot)
                LigneLot = Trim(ContenuLot(i))
                If Len(LigneLot) > 0 Then
                    ChampsLot = Split(LigneLot, ";")
                    If UBound(ChampsLot) >= 2 Then
                        Select Case ChampsLot(0)
                            Case "DB"  ' Début de barre
                                ProfilLot = Replace(ChampsLot(1), """", "")
                                Couleurs = Replace(ChampsLot(2), """", "")
                                Longueur = Replace(ChampsLot(3), """", "")
                                ClefBarre = Dictionnaire_SeqLotsBC(LotBC) & "-" & ProfilLot & "$" & ChampsLot(2) & "$" & ChampsLot(3) & "-" & LotCU
                                ' Incrémentation du compteur pour la barre
                                If Not Dictionnaire_BarresCU.Exists(ClefBarre) Then
                                    Dictionnaire_BarresCU(ClefBarre) = 1
                                Else
                                    Dictionnaire_BarresCU(ClefBarre) = Dictionnaire_BarresCU(ClefBarre) + 1
                                End If
                            Case "DP"  ' Début de pièce
                                ' Incrémentation du compteur pour la piècE
                                If Not Dictionnaire_PiecesCU.Exists(ProfilLot) Then
                                    Dictionnaire_PiecesCU(ProfilLot) = 1
                                Else
                                    Dictionnaire_PiecesCU(ProfilLot) = Dictionnaire_PiecesCU(ProfilLot) + 1
                                End If
                            Case "ET"
                                'On lit la ligne ET qui corresond aux infos sur l'étiquette. Si le lot de l'étiquette diffère du Lot du fichier alors c'est une optimisation.
                                'On cherche à garder l'information des rack optim afin de prévenir quand un morceau est déjà dans le rack (On évite les MANQPROF, on ajout également une condition pour passer les lots avec rack optim sur le même centre.
                                If Not LotCU Like ChampsLot(10) & "*" Then
                                    'Dictionnaire_RACKOPTIM(LotDest) = LotSource
                                    Dictionnaire_RACKOPTIM(ChampsLot(10) & TypeDuLot) = LotCU
                                    'J'ajoute le type du lot pour le retrouver en lot CU
                                    Dictionnaire_MorcOPTIM(ChampsLot(3)) = ChampsLot(10) & TypeDuLot
                                End If
                            Case "OP"
                                If ChampsLot(1) = "668" Or ChampsLot(1) = "669" Then
                                    Dictionnaire_OtPE(LotCU) = True
                                End If
                        End Select
                    End If
                End If
            Next i
            
            ' Détermination du type de lot et vérification du regroupement
            Regroupement = TestRegroupement(TypeDuLot)
            
            ' Écriture des résultats en fonction du mode de regroupement
            Select Case Regroupement
                Case False
                    ' Marquage du lot CU pour le jour de production
                    BesoinSemaineLotCU(LotCU) = True
                    If Dictionnaire_PiecesCU.count > 0 Then
                        CheminCHARGELOT = "U:\06- Projet\04- Alu\42 - Passage sur CU V14\%TEMP%\COMPIL_" & oFichier.Name
                        Set FichierCHARGELOT = Manipulationfichiers.CreateTextFile(CheminCHARGELOT, True)
                        ' Écriture de chaque profil et de son compteur dans le fichier
                        For Each Profil In Dictionnaire_PiecesCU.Keys
                            FichierCHARGELOT.WriteLine Profil & ";" & Dictionnaire_PiecesCU(Profil)
                            Dictionnaire_ContenuLot(oFichier.Name & ";" & Profil) = Dictionnaire_PiecesCU(Profil)
                        Next Profil
                        FichierCHARGELOT.Close
                        Dictionnaire_PiecesCU.RemoveAll
                    End If
                Case True
                    CheminCHARGELOT = "U:\06- Projet\04- Alu\42 - Passage sur CU V14\%TEMP%\COMPIL_Regroupement" & TypeDuLot & ".LOT"
                    ' Ouvrir en mode ajout si le fichier existe, sinon création d'un nouveau fichier
                    If Manipulationfichiers.FileExists(CheminCHARGELOT) Then
                        Set FichierCHARGELOT = Manipulationfichiers.OpenTextFile(CheminCHARGELOT, 8)
                    Else
                        Set FichierCHARGELOT = Manipulationfichiers.CreateTextFile(CheminCHARGELOT, True)
                    End If
                    For Each Profil In Dictionnaire_PiecesCU.Keys
                        FichierCHARGELOT.WriteLine Profil & ";" & Dictionnaire_PiecesCU(Profil)
                        Dictionnaire_ContenuLot("Regroupement" & TypeDuLot & ";" & Profil) = Dictionnaire_PiecesCU(Profil)
                    Next Profil
                    If Not BesoinSemaineLotCU.Exists("Regroupement" & TypeDuLot) Then BesoinSemaineLotCU("Regroupement" & TypeDuLot) = True
                    'Format de la donnée: Lot= Séquence;NombreOuvrants;NombreDormants
                    If Not Dictionnaire_SynthèseProd.Exists("Regroupement" & TypeDuLot) Then
                        Dictionnaire_SynthèseProd("Regroupement" & TypeDuLot) = "100000;0;0"
                    End If
                    
                    Dictionnaire_Regroupement("Regroupement" & TypeDuLot) = ContenuBRUT
                    
                    FichierCHARGELOT.Close
                    Dictionnaire_PiecesCU.RemoveAll
            End Select
        Next oFichier
LotSuivant:
        ' Passage au lot suivant
    Next LotBC
    Debug.Print "Éclatement des lots terminé."
End Sub

Public Function TestRegroupement(TypeDuLot As String) As Boolean
    TestRegroupement = False
    Select Case TypeDuLot
        Case "PAM207"
            If Configurateur.PAM207Reg.Value = True Then TestRegroupement = True
        Case "PAM208"
            If Configurateur.PAM208Reg.Value = True Then TestRegroupement = True
        Case "PAM306"
            If Configurateur.PAM306Reg.Value = True Then TestRegroupement = True
        Case "PAM322"
            If Configurateur.PAM322Reg.Value = True Then TestRegroupement = True
        Case "PAM305"
            If Configurateur.PAM305Reg.Value = True Then TestRegroupement = True
        Case Else
            ' Aucun regroupement par défaut pour les autres types
            TestRegroupement = False
    End Select
End Function

' Fonction pour déterminer le type de lot à partir de son nom
' Paramètre :
'   - NomDuLot (Variant) : nom complet du lot
' Retourne :
'   - Le type de lot sous forme de chaîne de caractères
Public Function DeterminerTypeDeLot(NomDuLot As String) As String
    ' Liste des types de lots possibles
    TypeLOTS = Array("AUTO", "COL", "COLIS", "COP", "COPCOUL", "DOR", "FBC030", "OUV", "PAM208", "PAM207", _
                     "PAM305", "PAM306", "PAM322", "PARCLOSE", "Scie", "VR")
    ' Parcours des types de lots pour déterminer le type correspondant
    For i = LBound(TypeLOTS) To UBound(TypeLOTS)
        ' Vérifie si la fin du nom du lot correspond au type
        If Right(NomDuLot, Len(TypeLOTS(i))) = TypeLOTS(i) Then
            DeterminerTypeDeLot = TypeLOTS(i)
            Exit For
        End If
    Next i
End Function
