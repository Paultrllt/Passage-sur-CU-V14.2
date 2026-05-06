Attribute VB_Name = "F_EcritureLOG"
'Ce module sert à rédiger le fichier LOG qui sera utilisé afin de comprendre et critiquer les décisions du logiciel de répartition
'Public Sub InitLogFile()
   'Dim FileSystem As Object
    'Dim Chemin_Log As String
   ' Chemin_Log = "U:\06- Projet\04- Alu\42 - Passage sur CU V14\logs.txt" ' Chemin du fichier de log
   ' Set FileSystem = CreateObject("Scripting.FileSystemObject")
   ' Set LogFile = FileSystem.CreateTextFile(Chemin_Log, True) ' Ouvrir le fichier en mode ajout
   ' LogFile.WriteLine Now & " - Début de la session de log."
'End Sub

' Fonction pour ajouter des logs
'Public Sub Logmessage(message As String)
    'If LogFile Is Nothing Then InitLogFile ' Initialisation si nécessaire
    'LogFile.WriteLine Now & " - " & message
'End Sub

' Fonction pour fermer le fichier de log
'Public Sub CloseLogFile()
    'If Not LogFile Is Nothing Then
       ' LogFile.WriteLine Now & " - Fin de la session de log."
        'LogFile.Close
       ' Set LogFile = Nothing
    'End If
'End Sub
