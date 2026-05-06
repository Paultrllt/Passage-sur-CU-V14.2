Attribute VB_Name = "H_GenerationSynthèseProd"
Sub GenerationSynthèseProd()
    Dim FeuilleSUIVI As Worksheet
    Dim FeuilleBase As Worksheet
    Dim TableauSUIVI As ListObject
    Dim NouvelleLigne As ListRow
    Dim InfosLOTBC() As String
    Dim Séquence As String
    Dim NombreOT As Integer
    Dim NombreDT As Integer
    Dim LotCU As Variant
    Dim item As Variant
    Dim Existe As Boolean
    Dim ligneExistante As ListRow
    Dim Morceau As Variant
    Dim Niter As Long
    Dim TempsPassage As Double
    
    ' Déclaration des variables pour les nouvelles colonnes
    Dim ColonneSortieOptim As Integer, ColonneEntreeOptim As Integer
    Dim ColonneRegroupement As Integer, ColonneCUBis As Integer, ColonneTempsPassage As Integer
    Dim SortieOptimValue As String, EntreeOptimValue As String
    Dim keyOptim As Variant
    
    ' Références des feuilles et du tableau
    Set FeuilleBase = ThisWorkbook.Worksheets("BDD_SynthèseProd")
    Set TableauSUIVI = FeuilleBase.ListObjects("Synthèse_SUIVI_PROD")
    
    ' Effacer les anciennes données du tableau
    If Not TableauSUIVI.DataBodyRange Is Nothing Then
        TableauSUIVI.DataBodyRange.Delete
    End If
    
    ' Récupération des index des colonnes existantes
    ColonneSéquence = TableauSUIVI.ListColumns("N°").Index
    ColonneLotCU = TableauSUIVI.ListColumns("Nom de lot").Index
    ColonneQteOt = TableauSUIVI.ListColumns("QTE Ouvrant").Index
    ColonneQteDt = TableauSUIVI.ListColumns("QTE Dormant").Index
    ColonneCU = TableauSUIVI.ListColumns("CU").Index
    ColonneJOUR = TableauSUIVI.ListColumns("Jour").Index
    ColonneSemaineFab = TableauSUIVI.ListColumns("Semaine").Index
    ColonneAnnée = TableauSUIVI.ListColumns("Année").Index
    ColonneCUBis = TableauSUIVI.ListColumns("CU Bis").Index
    
    ' Récupération des index des nouvelles colonnes
    ColonneSortieOptim = TableauSUIVI.ListColumns("Envoi OPTIM").Index
    ColonneEntreeOptim = TableauSUIVI.ListColumns("Récupération OPTIM").Index
    ColonneTempsPassage = TableauSUIVI.ListColumns("Temps Passage").Index

    
    ' Boucle sur chaque jour de production
    Niter = 0
        For Each LotCU In BesoinSemaineLotCU
            Séquence = ""
            NombreOT = 0
            NombreDT = 0
            
            'TypeLot
            Dim typeLot As String
            typeLot = DeterminerTypeDeLot(CStr(LotCU))
            
            ' Récupération des informations du lot
            Dim LotBC As String
            InfosLOTBC = Split(Dictionnaire_SynthèseProd(LotCU), ";")
            For Each item In Dictionnaire_SynthèseProd
                If LotCU Like item & "*" Then
                    ' Format de la donnée : Lot = Séquence;NombreOuvrants;NombreDormants
                    InfosLOTBC = Split(Dictionnaire_SynthèseProd(item), ";")
                    LotBC = item
                    Exit For
                End If
            Next item
            
            TempsPassage = Dictionnaire_ChargeCU(LotCU & "$" & Attribution_LotCU(LotCU))
            
            If UBound(InfosLOTBC) > 0 Then
                Séquence = InfosLOTBC(0)
                
                If typeLot = "OUV" Then NombreOT = InfosLOTBC(1)
                If typeLot = "DOR" Then NombreDT = InfosLOTBC(2)
    
                ' Vérification si le lot existe déjà dans le tableau
                Existe = False
                For Each ligneExistante In TableauSUIVI.ListRows
                    If ligneExistante.Range.Cells(1, ColonneLotCU).Value = LotCU Then
                        Set NouvelleLigne = ligneExistante
                        Existe = True
                        Exit For
                    End If
                Next ligneExistante
                
                ' Si le lot n'existe pas, on ajoute une nouvelle ligne
                If Not Existe Then
                    Set NouvelleLigne = TableauSUIVI.ListRows.Add
                End If
                
                ' Si aucune affectation n'est enregistrée pour ce lot, on indique "Manuel"
                If Attribution_LotCU(LotCU) = "" Then Attribution_LotCU(LotCU) = "Manuel"
    
                ' --- Détermination des colonnes OPTIM ---
                ' Sortie OPTIM : si le lot est présent dans le dictionnaire ADeplacerVersRackOptim,
                ' cela signifie que le morceau optimisé (celui débité lors du passage) doit être envoyé par chariot.
                SortieOptimValue = ""
                For Each lotDEST In Dictionnaire_RACKOPTIM
                    If LotCU = Dictionnaire_RACKOPTIM(lotDEST) Then
                        If Attribution_LotCU.Exists(Dictionnaire_RACKOPTIM(lotDEST)) Then
                            If Attribution_LotCU(lotDEST) <> Attribution_LotCU(Dictionnaire_RACKOPTIM(lotDEST)) Then SortieOptimValue = lotDEST & " Vers " & Attribution_LotCU(lotDEST)
                        End If
                    End If
                Next lotDEST
                
                ' Entrée OPTIM : si ce lot apparaît en tant que lot Dest dans Dictionnaire_RACKOPTIM,
                ' alors un morceau optimisé doit être récupéré dans son rack.
                Dim NbMorcRackOptim As Integer
                EntreeOptimValue = ""
                NbMorcRackOptim = 0
                For Each Morceau In Dictionnaire_MorcOPTIM.Keys
                    If Dictionnaire_MorcOPTIM(Morceau) = LotCU Then
                        NbMorcRackOptim = NbMorcRackOptim + 1
                    End If
                Next Morceau
                If NbMorcRackOptim <> 0 Then EntreeOptimValue = NbMorcRackOptim & " Morceaux."
                
                
                'CU Bis (si lotCU OUV donc centre DOR et réciproquement)
                Dim CUBis As String
                Select Case typeLot
                Case "OUV"
                    If Attribution_LotCU.Exists(LotBC & "DOR") Then
                        CUBis = Attribution_LotCU(LotBC & "DOR")
                    Else
                        CUBis = ""
                    End If
                    
                Case "DOR"
                    If Attribution_LotCU.Exists(LotBC & "OUV") Then
                        CUBis = Attribution_LotCU(LotBC & "OUV")
                    Else
                        CUBis = ""
                    End If
                    
                Case Else
                    CUBis = ""
                End Select
                
                ' Mise à jour (ou ajout) des informations dans la ligne
                With NouvelleLigne.Range
                    .Cells(1, ColonneSéquence).Value = Séquence
                    .Cells(1, ColonneLotCU).Value = LotCU
                    .Cells(1, ColonneQteOt).Value = NombreOT
                    .Cells(1, ColonneQteDt).Value = NombreDT
                    .Cells(1, ColonneCU).Value = Attribution_LotCU(LotCU)
                    .Cells(1, ColonneJOUR).Value = JourProdLotCU(LotCU)
                    .Cells(1, ColonneSemaineFab).Value = SFAB
                    .Cells(1, ColonneAnnée).Value = Année
                    .Cells(1, ColonneCUBis).Value = CUBis
                    ' Nouvelles colonnes OPTIM
                    .Cells(1, ColonneSortieOptim).Value = SortieOptimValue
                    .Cells(1, ColonneEntreeOptim).Value = EntreeOptimValue
                    .Cells(1, ColonneTempsPassage).Value = TempsPassage
                End With
        
            End If

        Next LotCU
    
End Sub


