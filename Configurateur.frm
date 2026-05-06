VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configurateur 
   Caption         =   "CONFIGURATEUR"
   ClientHeight    =   4920
   ClientLeft      =   516
   ClientTop       =   1836
   ClientWidth     =   8400.001
   OleObjectBlob   =   "Configurateur.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Configurateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Lancement_Click()
Call FluxPrincipal
End Sub

Public Sub UserForm_Initialize()
    ' Fixer la taille du UserForm
    Me.Width = 400  ' Largeur en points
    Me.Height = 265 ' Hauteur en points
    
    Me.BorderStyle = fmBorderStyleFixedSingle
    
    Call Initialisation 'Sub d'initialisation des Dictionnaires et autres objets
    Call LectureExportBC
    Call InitialisationSFABetSemaineExp
    
    SFABlabel.Caption = SFAB
    SEXPElabel.Caption = SemaineExp
    
End Sub

Public Sub CalculTempsUtile()
    Dim Centres As Variant
    Dim Fermetures As Variant
    Dim Nuits As Variant
    Dim TRS As Variant
    Dim TempsOuverture_DEBITenHeures As Long
    Dim TempsOuverture_DebitParEquipe As Long
    Dim TempsArretsPlan As Long
    Dim i As Integer
    Dim centre As String
    Dim Fermeture As Long
    Dim Nuit As Long
    Dim TRSValue As Double
    
    ' Définir les centres d'usinage et les variables associées
    Centres = Array("DUBUS3", "DUBUS5", "DUBUS8", "SCHIRMER4")
    TRS = Array(Configurateur.TRSD3INPUT, Configurateur.TRSD5INPUT, Configurateur.TRSD8INPUT, Configurateur.TRSSCH4INPUT)

    TempsOuverture_DEBITenHeures = (Configurateur.TempsOuvertureINPUT.Value / 5) * NombreJOURS
    TempsOuverture_DebitParEquipe = (Configurateur.TempsOuvertureINPUT.Value / 5)

    TempsArretsPlan = ((((20 + 5 + 5 + 45) * 2) * (NombreJOURS - 1)) + ((20 + 5 + 5 + 120) * 2)) / 60 '(15 minutes de pause + 5 minutes de réunions + 5 minutes d'échauffements + 45 minutes de nettoyage (Hors vendredi)* 2 Equipes + (Temps de pauses et échauffements standard + 2heures de nettoyage le vendredi)*2 équipes

    ' Boucle sur chaque centre d'usinage pour calculer le temps utile
    For i = 0 To UBound(Centres)
        centre = Centres(i)
        Fermeture = 0
        Nuit = 0
        
        ' Fermetures préventives
        If GetControlValue(Configurateur, "LPREVD" & Mid(centre, 5, 1)) = 1 Then Fermeture = Fermeture + 1
        If GetControlValue(Configurateur, "MAPREVD" & Mid(centre, 5, 1)) = 1 Then Fermeture = Fermeture + 1
        If GetControlValue(Configurateur, "MERPREVD" & Mid(centre, 5, 1)) = 1 Then Fermeture = Fermeture + 1
        If GetControlValue(Configurateur, "JEUPREVD" & Mid(centre, 5, 1)) = 1 Then Fermeture = Fermeture + 1
        If GetControlValue(Configurateur, "VPREVD" & Mid(centre, 5, 1)) = 1 Then Fermeture = Fermeture + 1
        
        ' Nuits
        If GetControlValue(Configurateur, "LND" & Mid(centre, 5, 1)) = 1 Then Nuit = Nuit + 1
        If GetControlValue(Configurateur, "MAND" & Mid(centre, 5, 1)) = 1 Then Nuit = Nuit + 1
        If GetControlValue(Configurateur, "MERND" & Mid(centre, 5, 1)) = 1 Then Nuit = Nuit + 1
        If GetControlValue(Configurateur, "JEUND" & Mid(centre, 5, 1)) = 1 Then Nuit = Nuit + 1
        If GetControlValue(Configurateur, "VND" & Mid(centre, 5, 1)) = 1 Then Nuit = Nuit + 1
        
        ' Calcul du temps utile pour chaque centre
        TRSValue = TRS(i).Value
        TempsUtile = ((TempsOuverture_DEBITenHeures * 2) - TempsArretsPlan - (Fermeture * TempsOuverture_DebitParEquipe) + (Nuit * TempsOuverture_DebitParEquipe)) * TRSValue
        
        Select Case centre
            Case "DUBUS3"
                TempsUtile_DUBUS3 = TempsUtile
            Case "DUBUS5"
                TempsUtile_DUBUS5 = TempsUtile
            Case "DUBUS8"
                TempsUtile_DUBUS8 = TempsUtile
            Case "SCHIRMER4"
                TempsUtile_SCHIRMER4 = TempsUtile
        End Select
        
        ' Mise à jour de l'affichage
        SetControlValue Configurateur, "T" & centre & ".Caption", TempsUtile & "h"
    Next i
End Sub
Public Sub Maj_Progression(Optional currentIter As Long = 0, Optional totalIter As Long = 1, Optional message As String = "")
    Dim progression As Double
    Dim pourcentage As Long
    Dim largeurMax As Long
    
    ' Vérification pour éviter une division par zéro
    If totalIter = 0 Then Exit Sub
    
    ' Calcul du pourcentage de progression
    progression = currentIter / totalIter
    pourcentage = Round(progression * 100, 0)
    
    ' On suppose que "FrameProgression" est le conteneur de la barre (définissez-le dans votre UserForm)
    largeurMax = Me.BarreProgression.Width
    
    With Me
        ' Si un message est fourni, on l'affiche
        If message <> "" Then .ETAPE.Caption = message
        
        ' Mise à jour de la largeur de la barre de progression
        .BarreProgression.Width = progression * largeurMax
        
        ' Mise à jour du label de pourcentage
        .AVANCEMENT.Caption = pourcentage & "%"
    End With
    
    DoEvents  ' Permet de rafraîchir l'affichage
End Sub



Public Sub Maj_EtiquettesSemaine()

SFABlabel.Caption = SFAB
SEXPElabel.Caption = SemaineExp
ProgressionExportsBC.Caption = IndexSemaineBC & " Exports BC traités / " & SemainesBCaTraiter & " Exports."
DoEvents
End Sub
