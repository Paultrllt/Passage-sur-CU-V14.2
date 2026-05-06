Attribute VB_Name = "G_GenerationExports"
' Fonction générique pour créer un fichier d'export
' Paramètre :
'   - chemin (String) : le chemin complet où le fichier sera créé
' Retourne :
'   - Un objet représentant le fichier texte créé
Public Function CreerFichierExport(chemin As String) As Object
    Dim FileSystem As Object
    ' Création d'un objet FileSystem pour gérer les fichiers
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    ' Création et ouverture du fichier texte en écriture
    Set CreerFichierExport = FileSystem.CreateTextFile(chemin, True)
End Function

' Sous-programme pour l'exportation des données MES
Public Sub ExportMES()
    Dim item As Variant
    Dim ItemLot As Variant
    Dim Chemin_EXPORTMES As String
    Dim ItemLotSansExtension As String
    Dim LotsParType As Object
    Dim LotSansType As String
    Dim typeLot As String
    Dim Niter As Long

    ' Définit le chemin du fichier d'export MES
    Chemin_EXPORTMES = "V:\PRODUCTION\SECTEUR DEBIT PREPARATION\test\EXPORT_MES_S" & SFAB & "_2.csv"
    ' Création du fichier d'export
    Set Fichier_EXPORTMES = CreerFichierExport(Chemin_EXPORTMES)
    
    ' Initialisation des messages et variables pour le suivi
    EtapeMessage = "Génération du fichier d'export MES"
    Niter = 0
    Debug.Print "Lancement export MES."

    ' Parcours des lots dans le dictionnaire
    For Each item In Dictionnaire_LotsBC
        ' Initialisation d'un dictionnaire pour regrouper les lots par type
        Set LotsParType = CreateObject("Scripting.Dictionary")

        ' Parcours des lots attribués
        For Each ItemLot In Attribution_LotCU
            
            ItemLotSansExtension = ItemLot
            ' Détermine le type de lot
            typeLot = DeterminerTypeDeLot(ItemLotSansExtension)
            LotSansType = Left(ItemLotSansExtension, Len(ItemLot) - Len(typeLot))
            
            If LotSansType = item Then
                If Attribution_LotCU.Exists(ItemLot) Then
                    ' Ajoute le lot au dictionnaire par type
                    LotsParType(typeLot) = Attribution_LotCU(ItemLot)
                End If
            End If
            
        Next ItemLot
        
        Debug.Print "Le lot " & item & " a été décomposé en " & LotsParType.count & " Sous-Lots."
        
        If LotsParType.Exists("OUV") Or LotsParType.Exists("DOR") Then
            ' Écriture des données au fichier d'export
            Fichier_EXPORTMES.WriteLine Dictionnaire_SeqLotsBC(item) & ";" & item & ";;;" & _
                LotsParType("OUV") & ";" & LotsParType("DOR")
        End If

        ' Mise à jour de la progression
        Niter = Niter + 1
        If Niter Mod 10 = 0 Then Configurateur.Maj_Progression Niter, Dictionnaire_LotsBC.count, "Export MES:"
    Next item

    ' Fermeture du fichier
    Fichier_EXPORTMES.Close
End Sub

' Sous-programme pour l'exportation des attributions
Public Sub ExportsAttributions()
    Dim ItemLotSansExtension As String
    Dim Chemin_EXPORT As String
    Dim ItemLot As Object
    Dim typeLot As String
    Dim LotSansType As String
    Dim CommandeRepere As String
    Dim Profil  As String
    Dim Position  As String
    Dim LongMorceau As String
    Dim CU  As String

    ' Initialisation des messages et variables pour le suivi
    Niter = 0
    totalIter = oDossierALU.Files.count
    EtapeMessage = "Export des attributions et données profils."
    Debug.Print "Lancement export des attributions et données profils."
    
    If IndexSemaineBC = 1 Then
        Set Fichier_EXPORT = CreerFichierExport("U:\06- Projet\04- Alu\42 - Passage sur CU V14\02_FichiersResultats\ExportAttributions.csv")
    End If

    ' Parcours des fichiers du dossier ALU
    For Each ItemLot In oDossierALU.Files
        ' Vérifie si le fichier a une extension
        If InStr(ItemLot.Name, ".") <> 0 Then
            ' Détermine le type de lot et le nom sans extension
            ItemLotSansExtension = Left(ItemLot.Name, InStr(ItemLot.Name, ".") - 1)
            typeLot = DeterminerTypeDeLot(ItemLotSansExtension)
            LotSansType = Left(ItemLotSansExtension, Len(ItemLotSansExtension) - Len(typeLot))

            ' Vérifie si le lot existe dans le dictionnaire
            If Dictionnaire_LotsBC.Exists(LotSansType) Then
                ' Lecture et traitement du fichier de lot
                Set FichierLOT = Manipulationfichiers.OpenTextFile(ItemLot, 1)
                ContenuLot = Split(FichierLOT.readall, vbNewLine)
                For i = 0 To UBound(ContenuLot) - 1
                    LigneLot = ContenuLot(i)
                    ChampsLot = Split(LigneLot, ";")
                    
                    If LigneLot <> "" Then
                    ' Vérifie si la ligne correspond à un repère (ET)
                    If ChampsLot(0) = "ET" Then
                        CommandeRepere = ChampsLot(11)
                        IDMORCEAU = ChampsLot(3)
                        Profil = ChampsLot(12)
                        Position = ChampsLot(5)
                        LongMorceau = ChampsLot(14)
                        CU = Attribution_LotCU(ItemLotSansExtension)
                        ' Écrit les données au fichier d'export
                        Fichier_EXPORT.WriteLine LotSansType & ";" & CommandeRepere & ";" & IDMORCEAU & ";" & Profil & ";" & Position & ";" & LongMorceau & ";" & CU & ";" & SFAB
                    End If
                    End If
                Next i
            End If
        End If

        ' Mise à jour de la progression
        Niter = Niter + 1
        If Niter Mod 10 = 0 Then Configurateur.Maj_Progression
    Next ItemLot
End Sub

' Sous-programme pour l'exportation des lots par centre
Public Sub ExportsContenuLots()
    Dim item As Variant
    Dim Chemin_ExportContenuLots As String
    Dim Niter As Long

    ' Définit le chemin du fichier d'export MES
    Chemin_ExportContenuLots = "U:\06- Projet\04- Alu\42 - Passage sur CU V14\02_FichiersResultats\ExportContenuLots_S" & SFAB & ".csv"
    ' Création du fichier d'export
    Set Fichier_ExportContenuLots = CreerFichierExport(Chemin_ExportContenuLots)
    
    ' Initialisation des messages et variables pour le suivi
    EtapeMessage = "Génération du fichier d'export contenu lots."
    Niter = 0

        ' Parcours des lots attribués
        For Each item In Dictionnaire_ContenuLot
            ' Écriture des données au fichier d'export
            Fichier_ExportContenuLots.WriteLine item & ";" & Dictionnaire_ContenuLot(item)
        ' Mise à jour de la progression
        Niter = Niter + 1
        If Niter Mod 10 = 0 Then Configurateur.Maj_Progression Niter, Dictionnaire_ContenuLot.count, "Génération du fichier d'export contenu lots."
        Next item
        
    ' Fermeture du fichier
    Fichier_ExportContenuLots.Close
End Sub
