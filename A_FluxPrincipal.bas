Attribute VB_Name = "A_FluxPrincipal"
Public Sub Initialisation()
Debug.Print "Init."
'Objet de manipulation des fichiers
    Set Manipulationfichiers = CreateObject("Scripting.FileSystemObject")

'Dictionnaires
    Set Dictionnaire_LotsBC = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_SeqLotsBC = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_ChargeCU = CreateObject("Scripting.Dictionary")
    Set Attribution_LotCU = CreateObject("Scripting.Dictionary")
    Set Lots_SURCHARGE = CreateObject("Scripting.Dictionary")
    Set LotsTraites = CreateObject("Scripting.Dictionary")
    Set LotsPrioritairesCUunique = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_SynthèseProd = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_Civières = CreateObject("Scripting.Dictionary")
    Set BesoinSemaineLotCU = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_Regroupement = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_RACKOPTIM = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_MorcOPTIM = CreateObject("Scripting.Dictionary")
    Set Dictionnaire_FichierAvecExtension = CreateObject("Scripting.Dictionary")
    
End Sub

Public Sub FluxPrincipal()
Debug.Print "FluxPrincipal."

SemainesBCaTraiter = oDossierExportBC.Files.count

    If SemainesBCaTraiter > 1 Then
        If MsgBox("Plusieurs fichiers d'ExportBC on été trouvés dans le dossier. Souhaitez-vous tous les traiter?", vbYesNo) = vbNo Then
            MsgBox ("Veuillez ne conserver que le fichier d'export à traiter dans le répertoire et relancer le programme.")
            Exit Sub
        End If
    End If
    
    Call PurgeTemp
        
        'On boucle le traitement pour chaque fichier d'export BC (Pour rappel: un fichier d'esport BC correspond à un semaine de prod)
        For Each FichierExportBC In oDossierExportBC.Files
        Debug.Print "Traitement d'une nouvelle semaine BC."
            
            IndexSemaineBC = IndexSemaineBC + 1
                
                Call RecuperationBesoinBC 'Sub de récupération du besoin pour la semaine
                Call EclatementLots 'Sub qui récupère les profils contenu dans un LotCu
                Call JourPROD
                Call GestionDesTemps 'Sub qui regroupe des sous-subs. Ces subs permettent de calculer les données temporelles
                Call RepartitionCUOptimale
                Call ExportMES
                'Call ExportsAttributions
                'Call ExportsLotsParCentre
                'Call ExportsContenuLots
                Call GenerationSynthèseProd
                'Call AffecterLotsRackOPTIM
                Call MiseEnLot
                'Call RegroupementLots
                    
            Dictionnaire_LotsBC.RemoveAll
            Dictionnaire_SeqLotsBC.RemoveAll
            Dictionnaire_ChargeCU.RemoveAll
            Attribution_LotCU.RemoveAll
            
        Next FichierExportBC
        
        Call PurgeTemp
        
        MsgBox "Répartition terminée!"

End Sub

Sub PurgeTemp()
    Dim dossier As String
    Dim nomFichier As String
    
    dossier = "U:\06- Projet\04- Alu\42 - Passage sur CU V14\%TEMP%\"
    
    ' Vérifier si le dossier existe
    If Dir(dossier, vbDirectory) = "" Then
        MsgBox "Le dossier n'existe pas.", vbExclamation
        Exit Sub
    End If
    
    ' Récupérer le premier fichier du dossier
    nomFichier = Dir(dossier & "*.*")
    
    ' Vérifier si le dossier contient des fichiers
    If nomFichier = "" Then
        Exit Sub
    Else

        If MsgBox("TEMP n'est pas purgé, souhaitez-vous supprimer les fichiers ? Si vous ne le faites pas, les calculs des lots regroupés pourront être faussés.", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
            ' Parcourir tous les fichiers du dossier
            Do While nomFichier <> ""
                ' S'assurer que l'élément n'est pas un sous-dossier
                If (GetAttr(dossier & nomFichier) And vbDirectory) <> vbDirectory Then
                    Kill dossier & nomFichier   ' Supprimer le fichier
                End If
                ' Passer au fichier suivant
                nomFichier = Dir
            Loop
        Else
            Exit Sub
        End If
    End If
End Sub

