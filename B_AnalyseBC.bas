Attribute VB_Name = "B_AnalyseBC"
' Ce module cherche à extraire le besoin de la semaine pour le répartir

' Ce sub récupère tous les fichiers d'exports BC dans le dossier
Public Sub LectureExportBC()
    Set oDossierExportBC = Manipulationfichiers.GetFolder("U:\06- Projet\04- Alu\42 - Passage sur CU V14\01_ExportBC")
End Sub

' Ce sub récupère les besoins à partir des fichiers d'export BC
Public Sub RecuperationBesoinBC()
    ' Déclaration des variables de traitement
    Dim Besoin As String
    Dim Lignes As Variant
    Dim i As Long
    Dim oFichier As Object
    Dim Champs As Variant
    Dim lot As String
    Dim SemaineExp As String
    
    ' Ouvrir le fichier d'export BC et lire son contenu
    Set oFichier = Manipulationfichiers.OpenTextFile(FichierExportBC, 1)
    Besoin = oFichier.readall
    oFichier.Close
    
    ' Diviser le contenu du fichier en lignes
    Lignes = Split(Besoin, vbNewLine)

    ' Parcourir chaque ligne du fichier
    For i = LBound(Lignes) To UBound(Lignes)
        If i Mod 10 = 0 Then Configurateur.Maj_Progression i, UBound(Lignes), "Récupération du besoin via export BC."
        
        ' Vérifier que la ligne n'est pas vide
        If Len(Trim(Lignes(i))) > 0 Then
            Champs = Split(Lignes(i), ";")

            ' Vérifier que les données sont valides
            If UBound(Champs) >= 7 And Champs(7) <> "Code lot" Then
                lot = Champs(7)
                ' Associer chaque lot à une date dans le dictionnaire
                Dictionnaire_LotsBC(lot) = CDate(Champs(1))
                ' Stocker les séquences dans un dictionnaire
                Dictionnaire_SeqLotsBC(lot) = Champs(2)
                ' Récupérer le numéro de la semaine d'expédition
                'Format de la donnée: Lot= Séquence;NombreOuvrants;NombreDormants
                Dictionnaire_SynthèseProd(lot) = Champs(2) & ";" & Champs(3) & ";" & Champs(4)
                SemaineExp = Champs(6)
            End If
        End If
    Next i
End Sub

' Ce sub détermine les jours de production pour chaque lot
Public Sub JourPROD()
    Dim TableDates() As Date
    Dim i As Long, j As Long
    Dim DatesUniques As Object
    Dim item As Variant
    Dim LotCU As Variant
    Dim jour(1 To 5) As Date
    Dim JourIndex As Long
    Dim LotsCuTraites As Object
    Dim DateElement As Variant
    Dim Index As Long

    ' Utilisation de dictionnaires pour stocker les dates uniques
    Set DatesUniques = CreateObject("Scripting.Dictionary")
    Set LotsCuTraites = CreateObject("Scripting.Dictionary")
    Set JourProdLotCU = CreateObject("Scripting.Dictionary")

    ' Création d'une table de dates à partir des lots
    ReDim TableDates(1 To Dictionnaire_LotsBC.count)
    For i = 1 To Dictionnaire_LotsBC.count
        TableDates(i) = Dictionnaire_LotsBC.items()(i - 1)
    Next i

    ' Tri rapide des dates (QuickSort)
    TriRapide TableDates, LBound(TableDates), UBound(TableDates)

    ' Ajouter les dates uniques au dictionnaire
    For i = LBound(TableDates) To UBound(TableDates)
        If Not DatesUniques.Exists(CStr(TableDates(i))) Then
            DatesUniques.Add CStr(TableDates(i)), TableDates(i)
        End If
    Next i

    ' Remplir les variables Jour(1 à 5) avec les dates uniques
    Index = 1
    For Each DateElement In DatesUniques.items
        If Index <= 5 Then
            jour(Index) = DateElement
            Index = Index + 1
        End If
    Next DateElement

    ' Déterminer le numéro de semaine ISO pour le premier jour
    Année = Year(DatesUniques.items()(0))
    ' Déterminer le nombre de jours en fonction des dates uniques disponibles
    NombreJOURS = Application.WorksheetFunction.Min(5, DatesUniques.count)

    ' Vérifier que le nombre de jours est valide
    If NombreJOURS = 0 Then
        MsgBox ("Incohérence de dates (Nombre de jours = 0)")
        Exit Sub
    End If
    
    Dim Niter As Long
    Niter = 0
    ' Charger les données pour chaque jour
    For i = 1 To NombreJOURS
        EtapeMessage = "Analyse de la charge pour le jour: " & i & "."
        totalIter = Dictionnaire_LotsBC.count
        For Each item In Dictionnaire_LotsBC
            Niter = Niter + 1
            If Niter Mod 10 = 0 Then Configurateur.Maj_Progression Niter, Dictionnaire_LotsBC.count, "Analyse de la charge pour le jour: " & i & "."
            
            ' Vérifier si la date du lot correspond au jour en cours
            If Dictionnaire_LotsBC(item) = jour(i) Then
                For Each LotCU In BesoinSemaineLotCU
                    If LotCU Like item & "*" Then
                        ' Ajouter le lot CU à la charge du jour correspondant
                        JourProdLotCU(LotCU) = i
                    End If
                Next LotCU
            End If
        Next item
    Next i
    
    'Ajout des Regroupement à la J1
    For Each LotCU In BesoinSemaineLotCU
        If LotCU Like "Regroupement" & "*" Then
            JourProdLotCU(LotCU) = 1
        End If
    Next LotCU
    
    ' Mise à jour des étiquettes de la semaine
    Configurateur.Maj_EtiquettesSemaine
End Sub

' Fonction de tri rapide (QuickSort) pour trier les dates
Private Sub TriRapide(arr As Variant, low As Long, high As Long)
    Dim pivot As Date
    Dim i As Long, j As Long
    Dim temp As Date


    ' Vérifie si le sous-tableau a plus d'un élément
    If low < high Then
        ' Sélectionne le pivot comme l'élément du milieu
        pivot = arr((low + high) \ 2)
        i = low
        j = high
        Do While i <= j
            ' Trouve un élément à gauche du pivot qui est plus grand que le pivot
            Do While arr(i) < pivot
                i = i + 1
            Loop
            ' Trouve un élément à droite du pivot qui est plus petit que le pivot
            Do While arr(j) > pivot
                j = j - 1
            Loop
            ' Si les indices i et j ne se sont pas croisés, échange les éléments
            If i <= j Then
                ' Échange les éléments arr(i) et arr(j)
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                ' Avance les indices i et j
                i = i + 1
                j = j - 1
            End If
        Loop
        ' Applique récursivement le tri rapide aux sous-tableaux
        TriRapide arr, low, j
        TriRapide arr, i, high
    End If
End Sub

Public Sub InitialisationSFABetSemaineExp()
    ' Déclaration des variables de traitement
    Dim Besoin As String
    Dim Lignes As Variant
    Dim i As Long
    Dim oFichier As Object
    Dim Champs As Variant
    Dim lot As String
    Dim Fichier As Variant
    
    For Each Fichier In oDossierExportBC.Files
        ' Ouvrir le fichier d'export BC et lire son contenu
        Set oFichier = Manipulationfichiers.OpenTextFile(Fichier, 1)
        Besoin = oFichier.readall
        oFichier.Close
        Exit For
    Next Fichier
    
    
    'Diviser le contenu du fichier en lignes
    Lignes = Split(Besoin, vbNewLine)

    ' Initialiser l'étape de progression
    EtapeMessage = "Récupération du besoin via export BC."

    ' Parcourir chaque ligne du fichier
    For i = LBound(Lignes) To UBound(Lignes)
        ' Vérifier que la ligne n'est pas vide
        If Len(Trim(Lignes(i))) > 0 Then
            Champs = Split(Lignes(i), ";")

            ' Vérifier que les données sont valides
            If UBound(Champs) >= 7 And Champs(7) <> "Code lot" Then
                SemaineExp = Champs(6)
                SFAB = DatePart("ww", Champs(1), vbMonday, vbFirstFourDays)
                Exit Sub
            End If
        End If
    Next i
    
End Sub
