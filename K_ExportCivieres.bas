Attribute VB_Name = "K_ExportCivieres"
Option Explicit

Public Sub DéfinitionCivières()
    Dim FeuilleBase As Worksheet
    Dim FeuilleCiviere As Worksheet
    Dim TableauSUIVI As ListObject
    Dim TableauCIVIERE As ListObject
    Dim ligneExistante As ListRow
    Dim NouvelleLigneCIV As ListRow
    Dim Barre As Variant
    Dim InfosBarre() As String
    Dim LotCU As String
    Dim LotCUEnCours As String
    Dim Profil As String
    Dim typeLot As String
    Dim Gamme As String
    Dim jour As String
    Dim CU As String
    Dim NombreBarres As Variant
    Dim sSequence As String
    Dim posDash As Long

    Dim ColonneSéquence As Long
    Dim ColonneLotCU As Long
    Dim ColonneQteOt As Long
    Dim ColonneQteDt As Long
    Dim ColonneCU As Long
    Dim ColonneJOUR As Long
    Dim ColonneSemaineFab As Long
    Dim ColonneAnnée As Long

    Dim CivIdIndex As Long
    Dim CivProfilIndex As Long
    Dim CivNbBarresIndex As Long

    ' Références des feuilles et des tableaux
    Set FeuilleBase = ThisWorkbook.Worksheets("BDD_SynthèseProd")
    Set TableauSUIVI = FeuilleBase.ListObjects("Synthèse_SUIVI_PROD")
    
    Set FeuilleCiviere = ThisWorkbook.Worksheets("Rapport Civières")
    Set TableauCIVIERE = FeuilleCiviere.ListObjects("Rapport_Civières")
    
    ' Effacer les anciennes données du tableau de rapport
    If Not TableauCIVIERE.DataBodyRange Is Nothing Then
        TableauCIVIERE.DataBodyRange.Delete
    End If

    ' Récupération des index des colonnes du tableau de synthèse
    With TableauSUIVI
        ColonneSéquence = .ListColumns("N°").Index
        ColonneLotCU = .ListColumns("Nom de lot").Index
        ColonneQteOt = .ListColumns("QTE Ouvrant").Index
        ColonneQteDt = .ListColumns("QTE Dormant").Index
        ColonneCU = .ListColumns("CU").Index
        ColonneJOUR = .ListColumns("Jour").Index
        ColonneSemaineFab = .ListColumns("Semaine").Index
        ColonneAnnée = .ListColumns("Année").Index
    End With

    ' Récupération des index des colonnes du tableau de rapport civières
    With TableauCIVIERE
        CivIdIndex = .ListColumns("Id_Civière").Index
        CivProfilIndex = .ListColumns("Profil").Index
        CivNbBarresIndex = .ListColumns("Nombre de barres").Index
    End With

    ' Parcours des lignes du tableau de synthèse
    For Each ligneExistante In TableauSUIVI.ListRows
        sSequence = CStr(ligneExistante.Range.Cells(1, ColonneSéquence).Value)
        If Trim(sSequence) <> "" Then
            ' Extraction de la "Gamme" et du "Jour"
            Gamme = Left(ligneExistante.Range.Cells(1, ColonneLotCU).Value, 7)
            If Len(sSequence) > 5 Then
                jour = Left(sSequence, Len(sSequence) - 5)
            Else
                jour = sSequence
            End If

            CU = CStr(ligneExistante.Range.Cells(1, ColonneCU).Value)
            LotCUEnCours = CStr(ligneExistante.Range.Cells(1, ColonneLotCU).Value)
            typeLot = DeterminerTypeDeLot(LotCUEnCours)
           
            ' Parcours des éléments du dictionnaire (en supposant qu'il s'agit d'un objet Dictionary)
            For Each Barre In Dictionnaire_BarresCU.Keys
                InfosBarre = Split(CStr(Barre), "$")
                If UBound(InfosBarre) >= 2 Then
                    LotCU = InfosBarre(1)
                    posDash = InStr(InfosBarre(2), "-")
                    If posDash > 0 Then
                        Profil = Left(InfosBarre(2), posDash - 1)
                    Else
                        Profil = InfosBarre(2)
                    End If
                    NombreBarres = Dictionnaire_BarresCU(Barre)
               
                    If LotCUEnCours = LotCU Then
                        ' Ajout d'une nouvelle ligne dans le tableau de rapport
                        Set NouvelleLigneCIV = TableauCIVIERE.ListRows.Add
                        With NouvelleLigneCIV.Range
                            .Cells(1, CivIdIndex).Value = jour & "-" & Gamme & typeLot & "-" & CU
                            .Cells(1, CivProfilIndex).Value = Profil
                            .Cells(1, CivNbBarresIndex).Value = NombreBarres
                        End With
                    End If
                End If
            Next Barre
        End If
    Next ligneExistante
End Sub

