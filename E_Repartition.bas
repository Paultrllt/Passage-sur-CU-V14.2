Attribute VB_Name = "E_Repartition"
Option Explicit
' ==============================================================================
' Module de répartition des charges CU optimisé avec contrôle du lissage global
' ==============================================================================
'
' Ce module répartit les lots entre différents centres d'usinage en
' optimisant l’affectation en fonction des capacités disponibles, des temps de
' traitement, et des contraintes spécifiques (dont le rack optim).
'
' Les améliorations incluent :
'   - Une affectation initiale (basée sur une affectation "optimale" ou aléatoire)
'     respectant les contraintes de capacité.
'   - Un recuit simulé (SA) qui, directement lancé sur la solution initiale,
'     échange des lots entre centres en évaluant globalement l'impact sur :
'         * La charge moyenne et l'écart type (par rapport aux capacités)
'         * La dispersion (pour regrouper les lots d'une même gamme)
'         * Les pénalités liées aux contraintes spécifiques (rack optim,
'           mise en chariot, changement de cales)
'
' La fonction d’évaluation "CalculerPerformance" combine ces critères (avec une
' pondération forte sur la dispersion) pour guider le SA.

Public Sub RepartitionCUOptimale()
    Dim dictCapaciteCentres As Object         ' Capacité maximale (en secondes) de chaque centre
    Dim dictTempsTotalCentres As Object        ' Temps total déjà affecté dans chaque centre pour le jour courant
    Dim dictHistoriqueCharges As Object        ' Historique des ratios journaliers par centre
    Dim dictHistoriqueTempsChargés As Object    ' Historique du temps chargé quotidien par centre

    ' Création des dictionnaires
    Set dictCapaciteCentres = CreateObject("Scripting.Dictionary")
    Set dictTempsTotalCentres = CreateObject("Scripting.Dictionary")
    Set dictHistoriqueCharges = CreateObject("Scripting.Dictionary")
    Set dictHistoriqueTempsChargés = CreateObject("Scripting.Dictionary")
    
    
    ' Initialiser les capacités maximales et remettre à zéro les charges
    InitialiserCapacites dictCapaciteCentres, dictTempsTotalCentres
    
    ' Initialiser l'historique pour chaque centre
    Dim centre As Variant
    For Each centre In dictCapaciteCentres.Keys
        Dim collTemps As New Collection
        dictHistoriqueTempsChargés.Add centre, collTemps
    Next centre
    
    Debug.Print "Début de la répartition optimale des charges CU."
    

    Dim colLots As Collection, LotCU As Variant, centreOpt As String, coutLot As Double
    
        ' Affectation initiale pour chaque
        For Each LotCU In BesoinSemaineLotCU
            centreOpt = ObtenirCentreOptimalPourLot(LotCU, dictCapaciteCentres, dictTempsTotalCentres)
            If centreOpt <> "" Then
                If Not Attribution_LotCU.Exists(LotCU) Then
                    Attribution_LotCU.Add LotCU, centreOpt
                    coutLot = CalculerCoutLot(LotCU, centreOpt)
                    dictTempsTotalCentres(centreOpt) = dictTempsTotalCentres(centreOpt) + coutLot
                End If
            End If
        Next LotCU
        
        ' Optimisation par recuit simulé
        OptimiserGlobalSA_Integre dictCapaciteCentres, dictTempsTotalCentres

    ' Exporter le résumé hebdomadaire dans le tableau Excel "Synthèse_Charge"
    SyntheseFinale dictTempsTotalCentres, dictCapaciteCentres
End Sub

' ==============================================================================
' OptimiserGlobalSA_Integre : Procédure de recuit simulé intégrée qui optimise directement
' la solution (affectation des lots) en tenant compte des contraintes.
'
' Les échanges proposés portent sur l'affectation de deux lots à des centres différents.
' Pour chaque échange, on calcule :
'   - Le coût actuel de chaque lot dans son centre.
'   - Le coût si le lot était affecté à l'autre centre.
'   - Le temps de changement de cales (fonction ChangementCalesDUBUS3).
'   - Une pénalité fixe (PenaliteOptim) pour la mise en chariot.
'
' La fonction CalculerPerformance (qui combine la charge moyenne, l'écart type et
' la dispersion) est utilisée pour décider d'accepter ou non l'échange.
'
Private Sub OptimiserGlobalSA_Integre(dictCapaciteCentres As Object, dictTempsTotalCentres As Object)
    Dim T As Double: T = 1000         ' Température initiale
    Dim alpha As Double: alpha = 0.99   ' Facteur de refroidissement
    Dim iterMax As Long: iterMax = 1000000
    Dim perfCourante As Double
    perfCourante = CalculerPerformance(dictTempsTotalCentres, dictCapaciteCentres)
    
    Dim i As Long
    Dim lot1 As Variant, lot2 As Variant
    Dim centre1 As String, centre2 As String
    Dim cout1 As Double, cout2 As Double, nouveauCout1 As Double, nouveauCout2 As Double
    Dim chargeAncienne1 As Double, chargeAncienne2 As Double
    Dim tempsCales1 As Long, tempsCales2 As Long
    Dim nouvellePerf As Double, deltaPerf As Double
    
    For i = 1 To iterMax
        If i Mod 100 = 0 Then
            Call Configurateur.Maj_Progression(1 - CInt(T), iterMax, "SA")
        End If
        
        ' Sélection aléatoire de deux lots parmi ceux affectés (excluant ceux soumis aux contraintes rack optim)
        lot1 = SelectionnerLotAleatoire(Attribution_LotCU, Dictionnaire_RACKOPTIM)
        lot2 = SelectionnerLotAleatoire(Attribution_LotCU, Dictionnaire_RACKOPTIM)
        
        If lot1 <> "" And lot2 <> "" Then
            centre1 = Attribution_LotCU(lot1)
            centre2 = Attribution_LotCU(lot2)
            
            ' Si les deux lots sont déjà sur le même centre, aucun échange n'est pertinent
            If centre1 = centre2 Then GoTo SuiteIter
            
            ' Calculer les coûts actuels
            cout1 = CalculerCoutLot(lot1, centre1)
            cout2 = CalculerCoutLot(lot2, centre2)
            ' Calculer les coûts en cas d'échange
            nouveauCout1 = CalculerCoutLot(lot1, centre2)
            nouveauCout2 = CalculerCoutLot(lot2, centre1)
            
            ' Sauvegarder les charges actuelles pour pouvoir annuler l'échange si besoin
            chargeAncienne1 = dictTempsTotalCentres(centre1)
            chargeAncienne2 = dictTempsTotalCentres(centre2)
            
            ' Calculer le temps de changement de cales (fonction liée aux contraintes rack optim)
            If centre1 = "DUBUS3-2" Then tempsCales1 = ChangementCalesDUBUS3(CStr(lot1), CStr(lot2)) Else tempsCales1 = 0
            If centre2 = "DUBUS3-2" Then tempsCales2 = ChangementCalesDUBUS3(CStr(lot2), CStr(lot1)) Else tempsCales2 = 0
            
            ' Mise à jour des charges dans chaque centre :
            '   - On retire le coût initial du lot
            '   - On ajoute le coût du lot dans le centre d'échange
            '   - On ajoute le temps de changement de cales
            '   - On ajoute la pénalité fixe pour la mise en chariot
            dictTempsTotalCentres(centre1) = chargeAncienne1 - cout1 + nouveauCout2 + tempsCales1 + ChargeRackOptim(lot2, centre1)
            dictTempsTotalCentres(centre2) = chargeAncienne2 - cout2 + nouveauCout1 + tempsCales2 + ChargeRackOptim(lot1, centre2)
            
            ' Mise à jour temporaire de l'affectation
            Attribution_LotCU(lot1) = centre2
            Attribution_LotCU(lot2) = centre1
            
            ' Calculer la nouvelle performance globale (fonction qui intègre moyenne, écart type et dispersion)
            nouvellePerf = CalculerPerformance(dictTempsTotalCentres, dictCapaciteCentres)
            deltaPerf = nouvellePerf - perfCourante
            
            ' Critère du recuit simulé :
            '   Si l'échange améliore la performance (deltaPerf négatif) ou est accepté
            '   selon la probabilité exp(-deltaPerf/T), alors on conserve la nouvelle solution.
            If deltaPerf < 0 Or Rnd() < Exp(-deltaPerf / T) Then
                perfCourante = nouvellePerf
            Else
                ' Sinon, on annule l'échange : restauration de l'affectation et des charges
                Attribution_LotCU(lot1) = centre1
                Attribution_LotCU(lot2) = centre2
                dictTempsTotalCentres(centre1) = chargeAncienne1
                dictTempsTotalCentres(centre2) = chargeAncienne2
            End If
        End If
SuiteIter:
        T = T * alpha
        If T < 0.01 Then Exit For
    Next i
End Sub

' ==============================================================================
' CalculerPerformance : Calcule une mesure de performance globale qui combine :
'   - La charge moyenne (ratio charge/capacité)
'   - L'écart type des ratios de charge
'   - La dispersion (mesure du regroupement des lots par gamme)
'
Private Function CalculerPerformance(dictTempsTotalCentres As Object, dictCapaciteCentres As Object) As Double
    Dim centre As Variant, ratio As Double
    Dim somme As Double, sommeCarre As Double
    Dim nbCentres As Integer
    Dim moyenne As Double, variance As Double, ecartType As Double
    Dim TempsDispo As Long
    
    nbCentres = 0: somme = 0: sommeCarre = 0
    For Each centre In dictTempsTotalCentres.Keys
        Dim chargeEffectif As Double
        chargeEffectif = dictTempsTotalCentres(centre)
        TempsDispo = dictCapaciteCentres(centre)
        ratio = chargeEffectif / TempsDispo
        somme = somme + ratio
        sommeCarre = sommeCarre + ratio * ratio
        nbCentres = nbCentres + 1
    Next centre
    
    If nbCentres > 0 Then
        moyenne = somme / nbCentres
        variance = (sommeCarre / nbCentres) - (moyenne ^ 2)
        If variance < 0 Then variance = 0
        ecartType = Sqr(variance)
    Else
        moyenne = 0: ecartType = 0
    End If
    
    ' Calculer le facteur de dispersion qui mesure le regroupement par gamme
    Dim facteurDispersion As Double
    facteurDispersion = MesurerDispersion
    
    ' Pondérations : on accorde un poids fort à la dispersion (ici 10)
    Dim poidsMoyenne As Double: poidsMoyenne = 2
    Dim poidsEcart As Double: poidsEcart = 1
    Dim poidsDispersion As Double: poidsDispersion = 1
    
    CalculerPerformance = poidsMoyenne * moyenne + poidsEcart * ecartType + poidsDispersion * facteurDispersion
End Function

' ==============================================================================
' CalculerCoutLot : Retourne le coût (temps de traitement) pour affecter un lot à un centre.
Private Function CalculerCoutLot(LotCU As Variant, centre As String) As Double
    Dim cle As String
    cle = LotCU & "$" & centre
    If Dictionnaire_ChargeCU.Exists(cle) Then
        CalculerCoutLot = Dictionnaire_ChargeCU(cle)
    Else
        CalculerCoutLot = 1E+99
    End If
End Function

' ==============================================================================
' MajChargeCentre : Met à jour la charge cumulée d'un centre après affectation d'un lot.
Private Sub MajChargeCentre(centre As String, coutLot As Double, dictCapaciteCentres As Object, _
                             dictTempsTotalCentres As Object)
    dictTempsTotalCentres(centre) = dictTempsTotalCentres(centre) + coutLot
End Sub

' ==============================================================================
' Adaptation de SyntheseFinale : Exporte dans le tableau Excel "Synthèse_Charge"
' le résumé hebdomadaire pour chaque centre en calculant :
'   - Le ratio moyen hebdomadaire (moyenne des ratios journaliers)
'   - Le temps total chargé en heures (somme des charges journalières)
Private Sub SyntheseFinale(dictTempsTotalCentres As Object, dictCapaciteCentres As Object)
    Dim Feuille As Worksheet
    Dim Tableau As ListObject
    Dim NouvelleLigne As ListRow
    Dim colCentre As Integer, colPourcentage As Integer, colTemps As Integer
    Const SEC_PAR_HEURE As Long = 3600
    
    Set Feuille = ThisWorkbook.Worksheets("Tableau de bord")
    Set Tableau = Feuille.ListObjects("Synthèse_Charge")
    
    If Not Tableau.DataBodyRange Is Nothing Then
        Tableau.DataBodyRange.Delete
    End If
    
    colCentre = Tableau.ListColumns("Centre").Index
    colPourcentage = Tableau.ListColumns("% De charge").Index
    colTemps = Tableau.ListColumns("Temps chargé").Index
    
    Dim centre As Variant, collHistorique As Collection, collTemps As Collection
    Dim ratioSemaine As Double
    
    For Each centre In dictCapaciteCentres.Keys

        ratioSemaine = dictTempsTotalCentres(centre) / dictCapaciteCentres(centre)
        
        Set NouvelleLigne = Tableau.ListRows.Add
        With NouvelleLigne.Range
            .Cells(1, colCentre).Value = centre
            .Cells(1, colPourcentage).Value = Format(ratioSemaine * 100, "0.00") & "%"
            .Cells(1, colTemps).Value = dictTempsTotalCentres(centre) / SEC_PAR_HEURE
        End With
    Next centre
End Sub
' ==============================================================================
' ChangerCalesDUBUS3 : Calcule la pénalité (en secondes) de changement de cales pour un lot
' passant sur le centre DUBUS3.
Private Function ChangementCalesDUBUS3(lotTeste As String, lotEchange As String) As Double
    Dim dictTypes As Object, dictTestCales As Object, lot As Variant, typeLot As String
    Set dictTestCales = CreateObject("Scripting.Dictionary")
    
    ' Constitution d'un dictionnaire pour lister les lots sur DUBUS3 selon leur type
    For Each lot In Attribution_LotCU
        If Attribution_LotCU(lot) = "DUBUS3-2" Then
            Select Case lot
                Case lotTeste
                    ' Ne pas ajouter, car on souhaite le retirer
                Case lotEchange
                    dictTestCales(lot) = DeterminerTypeDeLot(CStr(lot))
                Case Else
                    dictTestCales(lot) = DeterminerTypeDeLot(CStr(lot))
            End Select
        End If
    Next lot
    
    ' Si le lot testé n'était pas déjà sur DUBUS3, on l'ajoute artificiellement
    If Not dictTestCales.Exists(lotEchange) Then
        dictTestCales(lotEchange) = DeterminerTypeDeLot(CStr(lotEchange))
    End If
    
    ' Comptage des types distincts pour déterminer la pénalité de changement de cales
    Set dictTypes = CreateObject("Scripting.Dictionary")
    For Each lot In dictTestCales
        typeLot = dictTestCales(lot)
        If Not dictTypes.Exists(typeLot) And typeLot <> "PAM207" And typeLot <> "PAM208" And typeLot <> "" Then
            dictTypes.Add typeLot, 1
        End If
    Next lot
    
    ' Chaque type distinct supplémentaire ajoute 2700 secondes (45 minutes)
    ChangementCalesDUBUS3 = (dictTypes.count - 1) * 2700
End Function

' ==============================================================================
' ChargeRackOptim : Vérifie si un lot soumis à la contrainte rack optim est affecté
' au bon centre. Seuls les lots figurant dans Dictionnaire_RACKOPTIM sont concernés.
' Si le centre d'affectation du lot source (obtenu via Dictionnaire_RACKOPTIM)
' est différent du centre passé en paramètre
' est appliquée.
'
Private Function ChargeRackOptim(LotCU As Variant, centre As String) As Long
    Dim item As Variant
        ' Vérifier si le lot courant correspond (via le motif) à la clé actuelle
        If Dictionnaire_RACKOPTIM.Exists(LotCU) Then
            ' Si le lot source associé est affecté à un centre différent du centre passé,
            ' alors appliquer la pénalité
            If Attribution_LotCU(Dictionnaire_RACKOPTIM(LotCU)) <> centre Then
                Dim NbMorcRackOptim As Integer
                Dim Morceau As Variant
                NbMorcRackOptim = 0
                For Each Morceau In Dictionnaire_MorcOPTIM.Keys
                    If Dictionnaire_MorcOPTIM(Morceau) = LotCU Then
                        NbMorcRackOptim = NbMorcRackOptim + 1
                    End If
                Next Morceau
                ChargeRackOptim = 100 * NbMorcRackOptim
                Exit Function
            End If
        End If
    ChargeRackOptim = 0  ' Aucune pénalité à appliquer si aucune condition n'est remplie
End Function


' ==============================================================================
' MesurerDispersion : Calcule la dispersion (indice de répartition) des lots par gamme.
' Plus la dispersion est faible, mieux tous les lots d'une même gamme sont regroupés.
Private Function MesurerDispersion() As Double
    Dim sousDict As Object
    ' Définir la liste des centres disponibles
    Dim tabCentres As Variant
    tabCentres = Array("DUBUS3", "DUBUS5", "DUBUS8", "SCHIRM4")
    
    Dim dictCentres As Object
    Set dictCentres = CreateObject("Scripting.Dictionary")
    Dim centre As Variant
    For Each centre In tabCentres
        Set sousDict = CreateObject("Scripting.Dictionary")
        dictCentres.Add centre, sousDict
    Next centre
    
    ' Parcourir chaque lot et accumuler le nombre d'affectations par gamme
    Dim lotCourant As Variant, cleGamme As String, affectation As String, cleCentre As String, nomLot As String
    For Each lotCourant In Attribution_LotCU.Keys
        ' Normaliser le nom du lot
        nomLot = NormaliserNomLot(CStr(lotCourant))
        ' Construire la clé gamme : "J" + jour + les 7 premiers caractères du nom normalisé + type de lot
        cleGamme = "J" & JourProdLotCU(lotCourant) & Left(nomLot, 7) & DeterminerTypeDeLot(CStr(lotCourant))
        affectation = Attribution_LotCU(lotCourant)
        ' Extraire le nom du centre (avant le tiret)
        cleCentre = Split(affectation, "-")(0)
        If dictCentres.Exists(cleCentre) Then
            Set sousDict = dictCentres(cleCentre)
            If Not sousDict.Exists(cleGamme) Then
                sousDict.Add cleGamme, 1
            Else
                sousDict(cleGamme) = sousDict(cleGamme) + 1
            End If
        End If
    Next lotCourant
    
    ' Agréger les comptes de chaque gamme sur tous les centres
    Dim dictGammeGlobal As Object
    Set dictGammeGlobal = CreateObject("Scripting.Dictionary")
    Dim cleCentreCourant As Variant, cleGammeCourante As Variant
    For Each cleCentreCourant In dictCentres.Keys
        Set sousDict = dictCentres(cleCentreCourant)
        For Each cleGammeCourante In sousDict.Keys
            If Not dictGammeGlobal.Exists(cleGammeCourante) Then
                dictGammeGlobal.Add cleGammeCourante, sousDict(cleGammeCourante)
            Else
                dictGammeGlobal(cleGammeCourante) = dictGammeGlobal(cleGammeCourante) + sousDict(cleGammeCourante)
            End If
        Next cleGammeCourante
    Next cleCentreCourant
    
    Dim dispersionTotale As Double, nbGammes As Long
    dispersionTotale = 0: nbGammes = 0
    Dim totalEffectif As Long, sommeDiff As Double, indiceDispersion As Double
    Dim effectifs() As Variant, i As Long, j As Long, dictRepartition As Object
    
    For Each cleGammeCourante In dictGammeGlobal.Keys
        Set dictRepartition = CreateObject("Scripting.Dictionary")
        totalEffectif = 0
        For Each cleCentreCourant In dictCentres.Keys
            Set sousDict = dictCentres(cleCentreCourant)
            If sousDict.Exists(cleGammeCourante) Then
                dictRepartition.Add cleCentreCourant, sousDict(cleGammeCourante)
                totalEffectif = totalEffectif + sousDict(cleGammeCourante)
            End If
        Next cleCentreCourant
        
        If totalEffectif > 0 And dictRepartition.count > 1 Then
            sommeDiff = 0
            effectifs = dictRepartition.items
            For i = LBound(effectifs) To UBound(effectifs)
                For j = LBound(effectifs) To UBound(effectifs)
                    sommeDiff = sommeDiff + Abs(effectifs(i) - effectifs(j))
                Next j
            Next i
            indiceDispersion = sommeDiff / (2 * dictRepartition.count * totalEffectif)
        Else
            indiceDispersion = 0
        End If
        
        dispersionTotale = dispersionTotale + indiceDispersion
        nbGammes = nbGammes + 1
    Next cleGammeCourante
    
    If nbGammes > 0 Then
        MesurerDispersion = dispersionTotale / nbGammes
    Else
        MesurerDispersion = 0
    End If
End Function

' ==============================================================================
' NormaliserNomLot : Normalise le nom d'un lot en remplaçant certains préfixes par "LAFR".
Private Function NormaliserNomLot(lot As String) As String
    Dim prefixes As Variant, prefixCourant As Variant
    prefixes = Array("LADF", "LADP", "LASP", "LACT", "LAFC")
    For Each prefixCourant In prefixes
        If Left(lot, Len(prefixCourant)) = prefixCourant Then
            NormaliserNomLot = "LAFR" & Mid(lot, Len(prefixCourant) + 1)
            Exit Function
        End If
    Next prefixCourant
    NormaliserNomLot = lot
End Function
' ==============================================================================
' InitialiserCapacites : Initialise les capacités maximales (en secondes) et
' remet à zéro le temps affecté pour chaque centre.
Private Sub InitialiserCapacites(dictCapaciteCentres As Object, dictTempsTotalCentres As Object)
    Const SEC_PAR_HEURE As Long = 3600
    ' Par exemple, les temps utiles pour chaque centre sont définis dans des variables globales
    dictCapaciteCentres.Add "DUBUS3-2", TempsUtile_DUBUS3 * SEC_PAR_HEURE
    dictCapaciteCentres.Add "DUBUS5-2", TempsUtile_DUBUS5 * SEC_PAR_HEURE
    dictCapaciteCentres.Add "DUBUS8-2", TempsUtile_DUBUS8 * SEC_PAR_HEURE
    dictCapaciteCentres.Add "SCHIRM4-2", TempsUtile_SCHIRMER4 * SEC_PAR_HEURE
    
    Dim centre As Variant
    For Each centre In dictCapaciteCentres.Keys
        dictTempsTotalCentres.Add centre, 0
        Debug.Print "Capacité pour " & centre & " : " & dictCapaciteCentres(centre) & " secondes."
    Next centre
End Sub

' ------------------------------------------------------------------
' ObtenirCentreOptimalPourLot : Recherche le centre optimal pour un lot parmi ceux disponibles.
'
Private Function ObtenirCentreOptimalPourLot(LotCU As Variant, CapaciteCentres As Object, TempsTotalCentres As Object) As String
    Dim dctCentresCompatibles As Object
    Set dctCentresCompatibles = CreateObject("Scripting.Dictionary")
    Dim cle As Variant, NomCentre As String
    Dim PosDelim As Integer
    
    For Each cle In Dictionnaire_ChargeCU.Keys
        PosDelim = InStr(cle, "$")
        If PosDelim > 0 Then
            If Left(cle, PosDelim - 1) = LotCU Then
                NomCentre = Mid(cle, PosDelim + 1)
                If CapaciteCentres.Exists(NomCentre) Then
                    dctCentresCompatibles.Add NomCentre, Dictionnaire_ChargeCU(cle)
                End If
            End If
        End If
    Next cle
    
    ObtenirCentreOptimalPourLot = TrouverCentreOptimal(dctCentresCompatibles, CapaciteCentres, TempsTotalCentres)
End Function

' ------------------------------------------------------------------
' TrouverCentreOptimal : Sélectionne le centre optimal parmi ceux compatibles en
' évaluant un score basé sur la charge cumulée et le coût d’ajout du lot (normalisé par la capacité).
'
Private Function TrouverCentreOptimal(dctCentresCompatibles As Object, CapaciteCentres As Object, TempsTotalCentres As Object) As String
    Dim centre As Variant
    Dim scoreOptimal As Double: scoreOptimal = 1000000000#
    Dim scoreActuel As Double
    Dim centreChoisi As String: centreChoisi = ""
    
    For Each centre In dctCentresCompatibles.Keys
        scoreActuel = (TempsTotalCentres(centre) + dctCentresCompatibles(centre)) / CapaciteCentres(centre)
        Debug.Print "Evaluation du centre " & centre & " : Score = " & scoreActuel
        If scoreActuel < scoreOptimal Then
            scoreOptimal = scoreActuel
            centreChoisi = centre
        End If
    Next centre
    
    If centreChoisi = "" Then
        Debug.Print "Aucun centre optimal trouvé parmi " & dctCentresCompatibles.count & " centres compatibles."
    Else
        Debug.Print "Centre optimal choisi : " & centreChoisi & " avec un score de " & scoreOptimal
    End If
    
    TrouverCentreOptimal = centreChoisi
End Function

' ------------------------------------------------------------------
' SelectionnerLotAleatoire : Sélectionne aléatoirement un lot dans Attribution_LotCU
' en excluant ceux soumis aux contraintes de rack optim (déjà enregistrés dans Dictionnaire_RACKOPTIM).
'
Private Function SelectionnerLotAleatoire(AttributionLot As Object, RackOptim As Object) As Variant
    Dim eligible() As Variant
    Dim key As Variant
    Dim count As Long
    count = 0
    For Each key In AttributionLot.Keys
        If Not RackOptim.Exists(key) Then
            count = count + 1
            ReDim Preserve eligible(1 To count)
            eligible(count) = key
        End If
    Next key
    If count > 0 Then
        Dim idx As Long
        Randomize
        idx = Int(count * Rnd) + 1
        SelectionnerLotAleatoire = eligible(idx)
    Else
        SelectionnerLotAleatoire = ""
    End If
End Function
