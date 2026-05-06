Attribute VB_Name = "I_GenerationFichierLotRegroup"
Public Sub CreationDossierSemaine(Semaine As String)
    Dim Centres() As Variant
    
    Centres = Array("DUBUS3-2", "DUBUS5-2", "DUBUS8-2", "SCHIRM4-2")
    For Each centre In Centres
        'MkDir ("S:\ALU\OPTIM DUBUS\00_Regroupement de lots (Source CU V14)\" & centre & "\S" & Semaine)
    Next centre
End Sub

Public Sub RegroupementLots()
    ' Déclaration des variables
    Dim ClefFichierLOT As Variant
    Dim InfosLot() As String
    Dim ClefRegroupement As String
    Dim CheminLotReg As String
    Dim FichierLotReg As Object
    Dim Dictionnaire_Contenus As Object
    Dim Centres As String
    Dim Semaine As String
    Dim Dictionnaire_Fichiers As Object
    Dim Fichier As Variant
    Dim nomLot As String
    Dim JourTraité As Integer
    Dim Regroupement As Boolean
    Dim TypeDuLot As String
    Dim JourCharge As Object
    
    Set Dictionnaire_Fichiers = CreateObject("Scripting.Dictionary")
    
    If SFAB < 10 Then Semaine = "0" & CStr(SFAB) Else Semaine = CStr(SFAB)
    
    Call CreationDossierSemaine(Semaine)
    
    Dim oDossierALU As Object
    Set oDossierALU = Manipulationfichiers.GetFolder("S:\ALU\OPTIM DUBUS")
    For Each oFichierALU In oDossierALU.Files
        Dictionnaire_Fichiers.Add oFichierALU.Name, oFichierALU
    Next oFichierALU
    Debug.Print "Chargement des fichiers terminé. Total fichiers: " & Dictionnaire_Fichiers.count
    
    ' Création d'un dictionnaire pour regrouper le contenu par type de regroupement
    Set Dictionnaire_Contenus = CreateObject("Scripting.Dictionary")
    
        For Each LotCU In BesoinSemaineLotCU
        Regroupement = False
        
        
        If LotCU Like "Regroupement" & "*" Then Regroupement = True
        TypeDuLot = DeterminerTypeDeLot(CStr(LotCU))
        
        Select Case Regroupement
        
            Case False
                'récupération du nom sans extension
                nomFichier = Dictionnaire_FichierAvecExtension(LotCU)
                If Attribution_LotCU(LotCU) <> "Manuel" Then
                ClefRegroupement = "J" & JourProdLotCU(LotCU) & "_" & Left(LotCU, 7) & "_" & TypeDuLot & "_" & Attribution_LotCU(LotCU)
                    
                    ' Vérification et initialisation si nécessaire
                    If Not Dictionnaire_Contenus.Exists(ClefRegroupement) Then
                        Dictionnaire_Contenus.Add ClefRegroupement, "" ' Initialisation avec une chaîne vide
                    End If
                    
                    ' Lecture du contenu du fichier
                    Dim Contenu As String
                    Dim FichierLOT As Object
                    Set FichierLOT = Manipulationfichiers.OpenTextFile(Dictionnaire_Fichiers(nomFichier), 1)
                    Contenu = FichierLOT.readall
                    FichierLOT.Close
                    
                    ' Ajout du contenu regroupé en évitant le saut de ligne au début
                    If Dictionnaire_Contenus(ClefRegroupement) = "" Then
                        Dictionnaire_Contenus(ClefRegroupement) = Contenu
                    Else
                        Dictionnaire_Contenus(ClefRegroupement) = Dictionnaire_Contenus(ClefRegroupement) & Contenu
                    End If
                End If
                    
            Case True
                    ClefRegroupement = "J" & JourProdLotCU(LotCU) & "_" & "Regroupement" & "_" & TypeDuLot & "_" & Attribution_LotCU(LotCU)
                    
                    
                    ' Ajout du contenu regroupé en évitant le saut de ligne au début
                    If Dictionnaire_Contenus(ClefRegroupement) = "" Then
                        Dictionnaire_Contenus(ClefRegroupement) = Dictionnaire_Regroupement(LotCU)
                    Else
                        Dictionnaire_Contenus(ClefRegroupement) = Dictionnaire_Contenus(ClefRegroupement) & Dictionnaire_Regroupement(LotCU)
                    End If
            End Select
            
        Dim Niter As Long
        Niter = Niter + 1
         
        Next LotCU
    
    Niter = 0
        ' Écriture dans les fichiers en une seule passe
        Dim ClefType As Variant
        Dim InfosClefType() As String
        Dim TypeProd As String
        Dim jour As String
        Dim centre As String
        
        'For Each ClefType In Dictionnaire_Contenus.Keys
            'InfosClefType = Split(ClefType, "_")
            'jour = InfosClefType(0)
            'centre = InfosClefType(3)
            'TypeProd = InfosClefType(1) & InfosClefType(2)
        
            ' Construction du chemin du fichier
            'CheminLotReg = "S:\ALU\OPTIM DUBUS\00_Regroupement de lots (Source CU V14)\" & centre & "\S" & Semaine & "\" & jour & "_" & TypeProd & ".LOT"
            
            ' Vérification de l'existence du fichier et ouverture en conséquence
            'If Manipulationfichiers.FileExists(CheminLotReg) Then
                Set FichierLotReg = Manipulationfichiers.OpenTextFile(CheminLotReg, 8)
            Else
                'Set FichierLotReg = Manipulationfichiers.CreateTextFile(CheminLotReg, True)
            End If
            
            ' Écriture du contenu et fermeture
            'FichierLotReg.WriteLine Dictionnaire_Contenus(ClefType)
            'FichierLotReg.Close
            
        Niter = Niter + 1
        If i Mod 10 = 0 Then Configurateur.Maj_Progression Niter, Dictionnaire_Contenus.count, "Regroupement des lots."
        
        Next ClefType
    
End Sub


