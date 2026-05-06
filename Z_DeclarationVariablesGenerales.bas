Attribute VB_Name = "Z_DeclarationVariablesGenerales"
Option Explicit
'Variables techniques:
Public Manipulationfichiers As Object
Public Fichier_EXPORT As Object
Public SemainesBCaTraiter As Long
Public IndexSemaineBC As Long

'Variables générales
Public SemaineExp As String 'Variable de la semaine d'expédition de la semaine de fab en cours de traitement
Public TempsUtile As Long
Public NombreJOURS As Long
Public SFAB As Integer
Public Année As Long

Public Temps_DUBUS3 As Object
Public Temps_DUBUS5 As Object
Public Temps_DUBUS8 As Object
Public Temps_SCHIRMER4 As Object

Public TempsUtile_DUBUS3 As Long
Public TempsUtile_DUBUS5 As Long
Public TempsUtile_DUBUS8 As Long
Public TempsUtile_SCHIRMER4 As Long

Public TempsTotalCentres As Object

Public LotsSansCompatibilités As Object
Public Dictionnaire_ContenuLot As Object
Public JourProdLotCU As Object

Public BesoinSemaineLotCU As Object
Public LotsPrioritairesCUunique As Object
Public LotsTraites As Object

'Variables du besoin BC
Public oDossierExportBC As Object
Public FichierExportBC As Variant
Public Dictionnaire_LotsBC As Object
Public Dictionnaire_SeqLotsBC As Object
Public Dictionnaire_ChargeCU As Object

'Suivi de la progression


'Dictionnaire attributions
Public Attribution_LotCU As Object
Public Lots_SURCHARGE As Object

'Variables pour écriture des Logs
Public LogFile As Object ' Fichier de log

'Variables pour stocker les fichiers traités
Public oDossierALU As Object

'Dictionnaire pour suivi prod
Public Dictionnaire_SynthèseProd As Object

Public Dictionnaire_BarresCU As Object

Public Dictionnaire_Regroupement As Object

Public Dictionnaire_RACKOPTIM As Object
Public Dictionnaire_MorcOPTIM As Object
Public Dictionnaire_FichierAvecExtension As Object
Public Dictionnaire_OtPE As Object

