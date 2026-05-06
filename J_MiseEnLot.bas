Attribute VB_Name = "J_MiseEnLot"
'Génération d'un fichier pour la mise en lot

' Sous-programme pour l'exportation des lots par centre
Public Sub MiseEnLot()
    Dim item As Variant
    Dim Chemin_MiseEnLot As String
    Dim InfosBarre() As String
    
    Dim Séquence As String, IDBarre As String, lot As String

    ' Définit le chemin du fichier d'export MES
    Chemin_MiseEnLot = "\\bouvet.local\donnees\public\PRODUCTION\SECTEUR DEBIT PREPARATION\1_BARRETTAGE\MISE_EN_LOT\2023\Fichiers Mise En Lot\MiseEnLot_S" & SFAB & ".csv"
    ' Création du fichier d'export
    Set Fichier_MiseEnLot = CreerFichierExport(Chemin_MiseEnLot)

    ' Initialisation des messages et variables pour le suivi
    EtapeMessage = "Génération du fichier d'ExportsLotsParCentre"
    Niter = 0
    totalIter = Attribution_LotCU.count
    Debug.Print "Lancement export attribution lots."

        ' Parcours des lots attribués
        For Each item In Dictionnaire_BarresCU
            InfosBarre = Split(item, "-")
            Séquence = InfosBarre(0)
            IDBarre = InfosBarre(1)
            lot = InfosBarre(2)
            
            ' Écriture des données au fichier d'export
            Fichier_MiseEnLot.WriteLine Séquence & ";" & IDBarre & ";" & Attribution_LotCU(lot)
                
        ' Mise à jour de la progression
        Niter = Niter + 1
        If Niter Mod 10 = 0 Then Configurateur.Maj_Progression
        Next item
            
            
    ' Fermeture du fichier
    Fichier_MiseEnLot.Close
End Sub
