Attribute VB_Name = "ImportCosmos"
' =====================================================================
'  Import des informations COSMOS dans le PV
'  1. Sur le ticket COSMOS : clic sur le favori "PV COSMOS"
'     puis bouton "Copier pour le PV"
'  2. Dans le PV ouvert : Ctrl + Maj + I  (ou Alt+F8 > Importer_Cosmos)
'  Les champs vides dans COSMOS ne sont pas modifiés dans le PV.
'  Ce module peut être placé dans le classeur de macros personnelles
'  (PERSONAL.XLSB) : il agit sur le PV actif, sans modifier le PV lui-même.
' =====================================================================
Option Explicit

' Optionnel : ton nom tel qu'il apparaît dans la liste "Technicien SPIE"
' (ex : "MARTINEZ Philippe"). Laisser "" pour ne pas le remplir.
Private Const TECHNICIEN As String = ""

Private mErreurs As String
Private mNb As Long

Public Sub Importer_Cosmos()
Attribute Importer_Cosmos.VB_ProcData.VB_Invoke_Func = "I\n14"
    Dim txt As String, d As Object, typ As String
    Dim ws As Worksheet, nomFeuille As String, f As Variant
    Dim rep As VbMsgBoxResult, wb As Workbook

    txt = LirePressePapier()
    If InStr(1, txt, "#PVCOSMOS", vbBinaryCompare) = 0 Then
        MsgBox "Le presse-papier ne contient pas de données COSMOS." & vbCrLf & vbCrLf & _
               "Sur la page du ticket COSMOS : clique sur le favori « PV COSMOS »," & vbCrLf & _
               "puis sur le bouton « Copier pour le PV », et relance l'import.", _
               vbExclamation, "Import COSMOS"
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    If Not FeuilleExiste(wb, "IPHONE") Or Not FeuilleExiste(wb, "Inst_PC") Or Not FeuilleExiste(wb, "Data") Then
        MsgBox "Ouvre d'abord le PV Excel (fichier avec les feuilles IPHONE, Inst_PC...)," & vbCrLf & _
               "puis relance l'import.", vbExclamation, "Import COSMOS"
        Exit Sub
    End If

    Set d = Decoder(txt)
    typ = UCase$(Lire(d, "TYPE"))
    Select Case typ
        Case "IPHONE": nomFeuille = "IPHONE"
        Case "INST_PC": nomFeuille = "Inst_PC"
        Case "REST_PC": nomFeuille = "Rest_PC"
        Case "REST_PC_IP": nomFeuille = "Rest_PC_IP"
        Case Else
            MsgBox "Type de PV inconnu : " & typ, vbExclamation, "Import COSMOS"
            Exit Sub
    End Select
    Set ws = wb.Sheets(nomFeuille)

    ' Le PV contient déjà un autre dossier ?
    If Len(Trim$(CStr(ws.Range("Q4").Value))) > 0 And _
       StrComp(Trim$(CStr(ws.Range("Q4").Value)), Lire(d, "DEM"), vbTextCompare) <> 0 Then
        rep = MsgBox("La feuille " & nomFeuille & " contient déjà le dossier " & ws.Range("Q4").Value & "." & vbCrLf & _
                     "Remplacer par " & Lire(d, "DEM") & " ?" & vbCrLf & vbCrLf & _
                     "(Pense à vider la feuille avec le bouton de remise à zéro si besoin.)", _
                     vbYesNo + vbQuestion, "Import COSMOS")
        If rep <> vbYes Then Exit Sub
    End If

    ' Afficher la bonne feuille (comme les boutons du menu)
    ws.Visible = xlSheetVisible
    ws.Activate
    For Each f In Array("Menu", "IPHONE", "Inst_PC", "Rest_PC", "Rest_PC_IP")
        If f <> nomFeuille And FeuilleExiste(wb, CStr(f)) Then wb.Sheets(f).Visible = xlSheetHidden
    Next f

    mErreurs = ""
    mNb = 0

    ' En-tête commun à tous les PV
    Ecrire ws, "Q4", Lire(d, "DEM")
    Ecrire ws, "E5", Lire(d, "NOM")
    Ecrire ws, "Q5", Lire(d, "PRENOM")
    Ecrire ws, "E6", Lire(d, "SITE")
    Ecrire ws, "Q6", Lire(d, "BATIMENT")
    Ecrire ws, "E7", Lire(d, "ETAGE")
    Ecrire ws, "Q7", Lire(d, "BUREAU")

    Select Case typ
        Case "IPHONE"
            Ecrire ws, "E10", Lire(d, "IP_INST_SN")
            Ecrire ws, "E11", ModeleListe(wb, Lire(d, "IP_INST_MODELE"), "Tableau8", "A15:A33")
            Ecrire ws, "E12", Lire(d, "IP_INST_IMEI")
            Ecrire ws, "Q10", Lire(d, "IP_REST_SN")
            Ecrire ws, "Q11", ModeleListe(wb, Lire(d, "IP_REST_MODELE"), "Tableau8", "A15:A33")
            Ecrire ws, "Q12", Lire(d, "IP_REST_IMEI")
            Ecrire ws, "I31", Lire(d, "NOM")
            Ecrire ws, "P31", Lire(d, "PRENOM")
            EcrireSiVide ws, "H33", TECHNICIEN

        Case "INST_PC", "REST_PC", "REST_PC_IP"
            Ecrire ws, "E10", Lire(d, "PC_MODELE")
            Ecrire ws, "E11", Lire(d, "PC_SN")
            Ecrire ws, "E12", Lire(d, "PC_NETBIOS")
            Ecrire ws, "F15", Lire(d, "STATION_SN")
            Ecrire ws, "Q10", Lire(d, "ECRAN1_SN")
            Ecrire ws, "Q12", Lire(d, "ECRAN2_SN")
            Ecrire ws, "I49", Lire(d, "NOM")
            Ecrire ws, "P49", Lire(d, "PRENOM")
            EcrireSiVide ws, "H51", TECHNICIEN
            If typ = "REST_PC_IP" Then
                Ecrire ws, "E23", ModeleListe(wb, Lire(d, "IP_REST_MODELE"), "Liste_IPhone", "U3:U21")
                Ecrire ws, "Q23", Lire(d, "IP_REST_IMEI")
                Ecrire ws, "E24", Lire(d, "IP_REST_SN")
            End If
    End Select

    If Len(mErreurs) > 0 Then
        MsgBox mNb & " champ(s) rempli(s) depuis " & Lire(d, "SCTASK") & "." & vbCrLf & vbCrLf & _
               "Cellules non modifiables (feuille protégée ?) :" & mErreurs & vbCrLf & vbCrLf & _
               "Saisis-les à la main.", vbExclamation, "Import COSMOS"
    Else
        MsgBox mNb & " champ(s) rempli(s) depuis " & Lire(d, "SCTASK") & " (" & Lire(d, "DEM") & ")." & vbCrLf & vbCrLf & _
               "Reste à compléter : états, accessoires, tests, accompagnement," & vbCrLf & _
               "et bâtiment / étage / bureau s'ils sont vides.", vbInformation, "Import COSMOS"
    End If
End Sub

Private Function FeuilleExiste(wb As Workbook, ByVal nom As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not sh Is Nothing
End Function

' --- Lecture du presse-papier (texte) ---
Private Function LirePressePapier() As String
    Dim o As Object, s As String
    On Error Resume Next
    Set o = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If Not o Is Nothing Then
        o.GetFromClipboard
        s = o.GetText(1)
    End If
    If Len(s) = 0 Then
        s = CreateObject("htmlfile").parentWindow.clipboardData.GetData("Text")
    End If
    On Error GoTo 0
    LirePressePapier = s
End Function

' --- Texte "CLE=valeur" -> dictionnaire ---
Private Function Decoder(ByVal txt As String) As Object
    Dim d As Object, lignes() As String, i As Long, p As Long, l As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    txt = Replace(txt, vbCr, "")
    lignes = Split(txt, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        l = lignes(i)
        p = InStr(1, l, "=")
        If p > 1 And Left$(l, 1) <> "#" Then d(Trim$(Left$(l, p - 1))) = Trim$(Mid$(l, p + 1))
    Next i
    Set Decoder = d
End Function

Private Function Lire(d As Object, ByVal cle As String) As String
    If d.Exists(cle) Then Lire = Trim$(CStr(d(cle)))
End Function

' --- Écrit une valeur texte (sans toucher aux champs si la valeur est vide) ---
Private Sub Ecrire(ws As Worksheet, ByVal adr As String, ByVal v As String)
    v = Trim$(v)
    If Len(v) = 0 Then Exit Sub
    On Error GoTo Erreur
    ' l'apostrophe force le texte : évite que les IMEI / n° de série
    ' soient transformés en nombres (3,58E+14)
    ws.Range(adr).Value = "'" & v
    mNb = mNb + 1
    Exit Sub
Erreur:
    mErreurs = mErreurs & vbCrLf & "  " & adr & " (" & v & ")"
End Sub

Private Sub EcrireSiVide(ws As Worksheet, ByVal adr As String, ByVal v As String)
    If Len(Trim$(v)) = 0 Then Exit Sub
    If Len(Trim$(CStr(ws.Range(adr).Value))) = 0 Then Ecrire ws, adr, v
End Sub

' --- Fait correspondre "APPLE iPhone 13 Pro" à la liste des modèles du PV ---
Private Function ModeleListe(wb As Workbook, ByVal v As String, ByVal nomTable As String, ByVal plageSecours As String) As String
    Dim rng As Range, c As Range, nv As String, nc As String
    Dim best As String, bestLen As Long, court As String, nbCourt As Long

    v = NettoyerModele(v)
    ModeleListe = v
    If Len(v) = 0 Then Exit Function

    On Error Resume Next
    Set rng = wb.Sheets("Data").ListObjects(nomTable).DataBodyRange
    If rng Is Nothing Then Set rng = wb.Sheets("Data").Range(plageSecours)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function

    nv = Normaliser(v)
    For Each c In rng.Cells
        nc = Normaliser(CStr(c.Value))
        If Len(nc) > 0 Then
            If nc = nv Then
                ModeleListe = CStr(c.Value)
                Exit Function
            End If
            ' "iphone13pro128go" commence par "iphone13pro" (retenu seulement
            ' s'il n'y a qu'un seul candidat : sinon "13 Pro Max" pourrait gagner)
            If Left$(nc, Len(nv)) = nv Then
                court = CStr(c.Value): nbCourt = nbCourt + 1
            End If
            ' "iphone11pro" commence par "iphone11pro" (liste plus courte)
            If Left$(nv, Len(nc)) = nc And Len(nc) > bestLen Then
                best = CStr(c.Value): bestLen = Len(nc)
            End If
        End If
    Next c
    If nbCourt = 1 Then
        ModeleListe = court
    ElseIf Len(best) > 0 And bestLen >= Len(nv) - 1 Then
        ModeleListe = best
    End If
End Function

Private Function NettoyerModele(ByVal v As String) As String
    Dim p As Long
    v = Trim$(v)
    p = InStr(1, v, " - ")
    If p > 0 Then v = Left$(v, p - 1)
    If UCase$(Left$(v, 6)) = "APPLE " Then v = Mid$(v, 7)
    NettoyerModele = Trim$(v)
End Function

Private Function Normaliser(ByVal v As String) As String
    v = LCase$(NettoyerModele(v))
    v = Replace(v, " ", "")
    v = Replace(v, "plus", "+")
    Normaliser = v
End Function
