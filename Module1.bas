Attribute VB_Name = "Module1"
Sub GC()
Dim conn As Object ' D�clarez la variable pour l'objet Connexion
    On Error GoTo Erreur ' Activez la gestion des erreurs
    
    ' Initialisez l'objet Connexion
    Set conn = CreateObject("ADODB.Connection")
    
    ' Essayez d'�tablir une connexion
    conn.Open "Driver={MySQL ODBC 9.1 Unicode Driver};Server=localhost;Database=BD_Gestion_de_Commandes;User=root;Password=adam123;"
    
    ' Si la connexion r�ussit
    MsgBox "Connexion r�ussie !"
    
    
Erreur:
    ' Capturez et affichez les erreurs
    MsgBox "Erreur : " & Err.Description, vbCritical
    If Not conn Is Nothing Then conn.Close ' Fermez la connexion si elle est ouverte
    Set conn = Nothing

End Sub
