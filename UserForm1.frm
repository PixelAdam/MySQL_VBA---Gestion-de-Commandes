VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6480
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11748
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim lastRow As Long
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQLClient As String
    Dim strSQLProduit As String
    Dim strSQLCommande As String
    Dim strSQLLigneCommande As String
    Dim prodPrix As Variant
    Dim quantite As Variant
    Dim commandeDate As String
    Dim clientID As Long
    Dim produitID As Long
    Dim commandeID As Long

    ' Validation des champs obligatoires
    If Trim(UserForm1.Tclientnom.Value) = "" Or _
       Trim(UserForm1.Tclientemail.Value) = "" Or _
       Trim(UserForm1.Tclienttel.Value) = "" Or _
       Trim(UserForm1.Tprodnom.Value) = "" Or _
       Trim(UserForm1.Tprodcat.Value) = "" Or _
       Trim(UserForm1.Tquantite.Value) = "" Or _
       Trim(UserForm1.Tdate.Value) = "" Then
        MsgBox "Veuillez remplir tous les champs obligatoires.", vbExclamation
        Exit Sub
    End If

    ' Validation et préparation des champs numériques
    If IsNumeric(UserForm1.Tprodprix.Value) And UserForm1.Tprodprix.Value <> "" Then
        prodPrix = UserForm1.Tprodprix.Value
    Else
        MsgBox "Veuillez entrer un prix valide.", vbExclamation
        Exit Sub
    End If

    If IsNumeric(UserForm1.Tquantite.Value) And UserForm1.Tquantite.Value <> "" Then
        quantite = UserForm1.Tquantite.Value
    Else
        MsgBox "Veuillez entrer une quantité valide.", vbExclamation
        Exit Sub
    End If

    commandeDate = Trim(UserForm1.Tdate.Value) ' Format date (assurez-vous qu'il soit valide)

    ' Ajout des données dans Excel
    Sheets("Feuil1").Activate
    With Sheets("Feuil1")
        lastRow = .Cells(.Rows.Count, 2).End(xlUp).Row + 1
        .Cells(lastRow, 2).Value = Trim(UserForm1.Tclientnom.Value)
        .Cells(lastRow, 3).Value = Trim(UserForm1.Tclientemail.Value)
        .Cells(lastRow, 4).Value = Trim(UserForm1.Tclienttel.Value)
        .Cells(lastRow, 6).Value = Trim(UserForm1.Tprodnom.Value)
        .Cells(lastRow, 7).Value = Trim(UserForm1.Tprodcat.Value)
        .Cells(lastRow, 8).Value = quantite ' Quantité dans la 8ème colonne
        .Cells(lastRow, 9).Value = prodPrix ' Prix dans la 9ème colonne
        .Cells(lastRow, 10).Value = commandeDate
    End With

    ' Connexion à la base MySQL
    On Error GoTo Erreur
    strConn = "Driver={MySQL ODBC 9.1 Unicode Driver};Server=localhost;Database=BD_Gestion_de_Commandes;User=root;Password=adam123;"
    Set conn = CreateObject("ADODB.Connection")
    conn.Open strConn

    If conn.State = 0 Then
        MsgBox "Échec de la connexion à MySQL.", vbCritical
        Exit Sub
    End If

    ' Transactions
    conn.BeginTrans

    ' Requête pour insérer le client
    strSQLClient = "INSERT INTO clients (nom_complet, email, telephone) VALUES ('" & _
                   Replace(Trim(UserForm1.Tclientnom.Value), "'", "''") & "', '" & _
                   Replace(Trim(UserForm1.Tclientemail.Value), "'", "''") & "', '" & _
                   Replace(Trim(UserForm1.Tclienttel.Value), "'", "''") & "')"
    Debug.Print "Requête client : " & strSQLClient
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = strSQLClient
    cmd.Execute

    ' Récupérer l'ID du client inséré
    Set rs = conn.Execute("SELECT LAST_INSERT_ID()")
    clientID = rs.Fields(0).Value

    ' Requête pour insérer le produit
    strSQLProduit = "INSERT INTO produits (nom, categorie, pric) VALUES ('" & _
                    Replace(Trim(UserForm1.Tprodnom.Value), "'", "''") & "', '" & _
                    Replace(Trim(UserForm1.Tprodcat.Value), "'", "''") & "', " & _
                    prodPrix & ")"
    Debug.Print "Requête produit : " & strSQLProduit
    cmd.CommandText = strSQLProduit
    cmd.Execute

    ' Récupérer l'ID du produit inséré
    Set rs = conn.Execute("SELECT LAST_INSERT_ID()")
    produitID = rs.Fields(0).Value

    ' Requête pour insérer la commande
    strSQLCommande = "INSERT INTO commandes (client_id, date_cammande) VALUES (" & clientID & ", '" & Replace(commandeDate, "'", "''") & "')"
    Debug.Print "Requête commande : " & strSQLCommande
    cmd.CommandText = strSQLCommande
    cmd.Execute

    ' Récupérer l'ID de la commande insérée
    Set rs = conn.Execute("SELECT LAST_INSERT_ID()")
    commandeID = rs.Fields(0).Value

    ' Requête pour insérer la ligne de commande
    strSQLLigneCommande = "INSERT INTO ligne_commandes (commande_id, produit_id, quantite) VALUES (" & commandeID & ", " & produitID & ", " & quantite & ")"
    Debug.Print "Requête ligne commande : " & strSQLLigneCommande
    cmd.CommandText = strSQLLigneCommande
    cmd.Execute

    ' Commit
    conn.CommitTrans
    MsgBox "Données insérées avec succès dans MySQL et Excel!", vbInformation

    ' Nettoyage
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    Exit Sub

Erreur:
    conn.RollbackTrans
    MsgBox "Erreur lors de l'insertion des données : " & Err.Description, vbCritical
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
    Set cmd = Nothing

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub Tprodnom_Change()

End Sub

Private Sub UserForm_Click()

End Sub
