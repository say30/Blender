Attribute VB_Name = "PVCosmosAuto"
' =====================================================================
'  PV COSMOS - création automatique des PV et de la charte iPhone
'
'  Fonctionnement :
'   1. Dans COSMOS, le favori "PV COSMOS" -> bouton "Créer les PV"
'      télécharge un petit fichier PVCOSMOS_<ticket>.txt
'   2. Ce module (dans PERSONAL.XLSB) surveille le dossier
'      Téléchargements toutes les 3 secondes tant qu'Excel est ouvert.
'   3. Pour chaque fichier trouvé : ouverture de PV_Vide.xlsm (et de la
'      charte), remplissage de l'onglet voulu, enregistrement dans
'        <COSMOS>\NOM Prénom\AAAA-MM-JJ\NOM Prénom - <opération>.xlsm
'      et charte en .xlsm + .pdf. Les modèles vides ne sont jamais modifiés.
'
'  Ctrl+Maj+I (à attribuer via Alt+F8 > Options) : traiter tout de suite
'  et (re)démarrer la surveillance.
' =====================================================================
Option Explicit

' ---- Réglages -------------------------------------------------------
' Dossier COSMOS. Vide = <Mes documents>\COSMOS
Private Const DOSSIER_COSMOS As String = ""
Private Const NOM_PV_VIDE As String = "PV_Vide.xlsm"
Private Const NOM_CHARTE_VIDE As String = "Charte_Iphone_Vide.xlsm"
' Nom tel qu'il apparaît dans la liste "Technicien SPIE" du PV
Private Const TECHNICIEN As String = "MARTINEZ Philippe"
Private Const INTERVALLE_SECONDES As Long = 3
' ---------------------------------------------------------------------

Private mProchain As Date
Private mActif As Boolean
Private mEnCours As Boolean
Private mErreurs As String
Private mNb As Long

' Lancé automatiquement à l'ouverture d'Excel (PERSONAL.XLSB)
Public Sub Auto_Open()
    DemarrerSurveillance
End Sub

Public Sub Auto_Close()
    ArreterSurveillance
End Sub

' Ctrl+Maj+I : traiter maintenant + démarrer la surveillance
Public Sub Importer_Cosmos()
    DemarrerSurveillance
    If Not mEnCours Then
        If TraiterTelechargements() = 0 Then
            MsgBox "Aucune demande en attente dans le dossier :" & vbCrLf & DossierTelechargements() & vbCrLf & vbCrLf & _
                   "Dans COSMOS : favori « PV COSMOS » puis « Créer les PV »." & vbCrLf & _
                   "La surveillance automatique est active tant qu'Excel reste ouvert.", vbInformation, "PV COSMOS"
        End If
    End If
End Sub

Public Sub DemarrerSurveillance()
    ArreterSurveillance
    mActif = True
    Programmer
End Sub

Public Sub ArreterSurveillance()
    mActif = False
    On Error Resume Next
    If mProchain <> 0 Then Application.OnTime mProchain, NomTic(), , False
    mProchain = 0
End Sub

Private Function NomTic() As String
    NomTic = "'" & ThisWorkbook.Name & "'!PVCosmos_Tic"
End Function

Private Sub Programmer()
    mProchain = Now + TimeSerial(0, 0, INTERVALLE_SECONDES)
    Application.OnTime mProchain, NomTic()
End Sub

Public Sub PVCosmos_Tic()
    On Error GoTo Fin
    If Not mActif Then Exit Sub
    If Not mEnCours Then TraiterTelechargements
Fin:
    mEnCours = False
    If mActif Then Programmer
End Sub

' ---------------------------------------------------------------------
'  Dossiers
' ---------------------------------------------------------------------
Private Function DossierTelechargements() As String
    Dim p As String
    On Error Resume Next
    p = CreateObject("Shell.Application").Namespace("shell:Downloads").Self.Path
    On Error GoTo 0
    If Len(p) = 0 Then p = Environ$("USERPROFILE") & "\Downloads"
    DossierTelechargements = p
End Function

Private Function DossierCosmos() As String
    Dim c As Variant, md As String
    If Len(DOSSIER_COSMOS) > 0 Then DossierCosmos = DOSSIER_COSMOS: Exit Function
    On Error Resume Next
    md = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    On Error GoTo 0
    ' Documents classique, Documents redirigé, Documents dans OneDrive
    For Each c In Array(Environ$("USERPROFILE") & "\Documents\COSMOS", md & "\COSMOS", _
                        Environ$("OneDrive") & "\Documents\COSMOS", Environ$("OneDriveCommercial") & "\Documents\COSMOS")
        If Len(Modele(CStr(c), "PV")) > 0 Then DossierCosmos = CStr(c): Exit Function
    Next c
    DossierCosmos = Environ$("USERPROFILE") & "\Documents\COSMOS"
End Function

' Trouve le modèle vide : nom exact, sinon "PV Vide.xlsm", "PV_Vide (1).xlsm"...
Private Function Modele(ByVal base As String, ByVal quoi As String) As String
    Dim f As String, motif As Variant
    If Len(base) < 4 Then Exit Function
    On Error Resume Next
    If quoi = "PV" Then
        If Len(Dir(base & "\" & NOM_PV_VIDE)) > 0 Then Modele = base & "\" & NOM_PV_VIDE: Exit Function
        motif = "PV*Vide*.xls*"
    Else
        If Len(Dir(base & "\" & NOM_CHARTE_VIDE)) > 0 Then Modele = base & "\" & NOM_CHARTE_VIDE: Exit Function
        motif = "Charte*Vide*.xls*"
    End If
    f = Dir(base & "\" & motif)
    If Len(f) > 0 Then Modele = base & "\" & f
End Function

' Ferme le modèle s'il est déjà ouvert dans Excel (sinon Excel refuse de l'ouvrir une 2e fois)
Private Sub FermerSiOuvert(ByVal chemin As String)
    Dim w As Workbook, nom As String
    nom = Mid$(chemin, InStrRev(chemin, "\") + 1)
    For Each w In Application.Workbooks
        If StrComp(w.Name, nom, vbTextCompare) = 0 Then
            If w.Saved Then
                w.Close SaveChanges:=False
            Else
                Err.Raise 1004, , "Le modèle " & nom & " est ouvert et modifié : ferme-le puis reclique sur « Créer les PV »."
            End If
        End If
    Next w
End Sub

Private Sub CreerDossiers(ByVal chemin As String)
    Dim parties() As String, i As Long, cur As String
    parties = Split(chemin, "\")
    cur = parties(0)
    For i = 1 To UBound(parties)
        cur = cur & "\" & parties(i)
        If Len(Dir(cur, vbDirectory)) = 0 Then MkDir cur
    Next i
End Sub

Private Function NomPropre(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, " ")
    Next c
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NomPropre = Trim$(s)
End Function

Private Function CheminUnique(ByVal dossier As String, ByVal nom As String, ByVal ext As String) As String
    Dim p As String, n As Long
    p = dossier & "\" & nom & ext
    n = 1
    Do While Len(Dir(p)) > 0
        n = n + 1
        p = dossier & "\" & nom & " (" & n & ")" & ext
    Loop
    CheminUnique = p
End Function

' ---------------------------------------------------------------------
'  Traitement des demandes téléchargées
' ---------------------------------------------------------------------
Private Function TraiterTelechargements() As Long
    Dim dl As String, f As String, liste As New Collection, it As Variant
    dl = DossierTelechargements()
    f = Dir(dl & "\PVCOSMOS_*.txt")
    Do While Len(f) > 0
        liste.Add dl & "\" & f
        f = Dir()
    Loop
    If liste.Count = 0 Then Exit Function
    mEnCours = True
    For Each it In liste
        TraiterFichier CStr(it)
    Next it
    mEnCours = False
    TraiterTelechargements = liste.Count
End Function

Private Sub TraiterFichier(ByVal fichier As String)
    Dim txt As String, d As Object, base As String, dossier As String
    Dim nomPers As String, actions() As String, i As Long, fait As String
    Dim dateInterv As Date, ok As Boolean

    txt = LireUtf8(fichier)
    On Error Resume Next
    Kill fichier              ' la demande est en mémoire : on la retire tout de suite
    On Error GoTo 0
    If InStr(1, txt, "#PVCOSMOS", vbBinaryCompare) = 0 Then Exit Sub
    Set d = Decoder(txt)

    base = DossierCosmos()
    If Len(Modele(base, "PV")) = 0 Then
        MsgBox "PV vide introuvable dans :" & vbCrLf & base & vbCrLf & vbCrLf & _
               "Place PV_Vide.xlsm (et Charte_Iphone_Vide.xlsm) dans ce dossier, puis reclique sur « Créer les PV » dans COSMOS.", _
               vbExclamation + vbSystemModal, "PV COSMOS"
        Exit Sub
    End If

    dateInterv = LireDate(Lire(d, "DATE"))
    nomPers = NomPropre(UCase$(Lire(d, "NOM")) & " " & Prenom(Lire(d, "PRENOM")))
    If Len(nomPers) = 0 Then nomPers = NomPropre(Lire(d, "DEM"))
    dossier = base & "\" & nomPers & "\" & Format(dateInterv, "yyyy-mm-dd")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo Erreur
    CreerDossiers dossier
    mErreurs = ""
    actions = Split(Lire(d, "ACTIONS"), ",")
    For i = LBound(actions) To UBound(actions)
        Select Case UCase$(Trim$(actions(i)))
            Case "INST_PC": fait = fait & vbCrLf & "  • " & CreerPV(d, base, dossier, "Inst_PC", nomPers & " - Installation PC")
            Case "REST_PC": fait = fait & vbCrLf & "  • " & CreerPV(d, base, dossier, "Rest_PC", nomPers & " - Désinstallation PC")
            Case "REST_PC_IP": fait = fait & vbCrLf & "  • " & CreerPV(d, base, dossier, "Rest_PC_IP", nomPers & " - Désinstallation PC + iPhone")
            Case "IPHONE": fait = fait & vbCrLf & "  • " & CreerPV(d, base, dossier, "IPHONE", nomPers & " - " & LibelleIphone(d))
            Case "CHARTE": fait = fait & vbCrLf & "  • " & CreerCharte(d, base, dossier, nomPers & " - Charte iPhone", dateInterv)
        End Select
    Next i
    ok = True
Erreur:
    If Not ok Then mErreurs = mErreurs & vbCrLf & "  Erreur : " & Err.Description
    On Error GoTo 0
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Journal base, d, dossier, fait
    On Error Resume Next
    Shell "explorer.exe """ & dossier & """", vbNormalFocus
    On Error GoTo 0
    MsgBox Lire(d, "TICKET") & " – " & nomPers & vbCrLf & vbCrLf & _
           "Fichiers créés dans :" & vbCrLf & dossier & fait & _
           IIf(Len(mErreurs) > 0, vbCrLf & vbCrLf & "À vérifier :" & mErreurs, "") & vbCrLf & vbCrLf & _
           "Reste à compléter le jour J : états, accessoires, tests, signatures.", _
           IIf(Len(mErreurs) > 0, vbExclamation, vbInformation) + vbSystemModal, "PV COSMOS"
End Sub

Private Function LibelleIphone(d As Object) As String
    Dim inst As Boolean, rest As Boolean
    inst = Len(Lire(d, "IPI_SN") & Lire(d, "IPI_IMEI")) > 0
    rest = Len(Lire(d, "IPR_SN") & Lire(d, "IPR_IMEI")) > 0
    If inst And rest Then
        LibelleIphone = "Changement iPhone"
    ElseIf rest Then
        LibelleIphone = "Restitution iPhone"
    Else
        LibelleIphone = "Installation iPhone"
    End If
End Function

' ---------------------------------------------------------------------
'  PV
' ---------------------------------------------------------------------
Private Function CreerPV(d As Object, ByVal base As String, ByVal dossier As String, _
                         ByVal feuille As String, ByVal nomFichier As String) As String
    Dim wb As Workbook, ws As Worksheet, f As Variant, chemin As String

    On Error GoTo Echec
    FermerSiOuvert Modele(base, "PV")
    Set wb = Workbooks.Open(Filename:=Modele(base, "PV"), UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set ws = wb.Sheets(feuille)
    ws.Visible = xlSheetVisible
    For Each f In Array("Menu", "IPHONE", "Inst_PC", "Rest_PC", "Rest_PC_IP")
        If f <> feuille Then wb.Sheets(f).Visible = xlSheetHidden
    Next f

    ' En-tête commun
    Ecrire ws, "Q4", Lire(d, "DEM")
    Ecrire ws, "E5", UCase$(Lire(d, "NOM"))
    Ecrire ws, "Q5", UCase$(Lire(d, "PRENOM"))
    Ecrire ws, "E6", Lire(d, "SITE")
    Ecrire ws, "Q6", Lire(d, "BATIMENT")
    Ecrire ws, "E7", Lire(d, "ETAGE")
    Ecrire ws, "Q7", Lire(d, "BUREAU")

    Select Case feuille
        Case "IPHONE"
            Ecrire ws, "E10", Lire(d, "IPI_SN")
            Ecrire ws, "E11", ModeleListe(wb, Lire(d, "IPI_MODELE"), "Data", "Tableau8", "A15:A33")
            Ecrire ws, "E12", Lire(d, "IPI_IMEI")
            Ecrire ws, "Q10", Lire(d, "IPR_SN")
            Ecrire ws, "Q11", ModeleListe(wb, Lire(d, "IPR_MODELE"), "Data", "Tableau8", "A15:A33")
            Ecrire ws, "Q12", Lire(d, "IPR_IMEI")
            Ecrire ws, "I31", UCase$(Lire(d, "NOM"))
            Ecrire ws, "P31", UCase$(Lire(d, "PRENOM"))
            Ecrire ws, "H33", TECHNICIEN
        Case "Inst_PC"
            RemplirPC ws, d, "PCI_"
        Case "Rest_PC"
            RemplirPC ws, d, "PCR_"
        Case "Rest_PC_IP"
            RemplirPC ws, d, "PCR_"
            Ecrire ws, "E23", ModeleListe(wb, Lire(d, "IPR_MODELE"), "Data", "Liste_IPhone", "U3:U21")
            Ecrire ws, "Q23", Lire(d, "IPR_IMEI")
            Ecrire ws, "E24", Lire(d, "IPR_SN")
    End Select

    AppliquerReglages wb, d, feuille

    wb.Activate
    ws.Activate
    chemin = CheminUnique(dossier, nomFichier, ".xlsm")
    wb.SaveAs Filename:=chemin, FileFormat:=52
    wb.Close SaveChanges:=False
    CreerPV = Mid$(chemin, InStrRev(chemin, "\") + 1)
    Exit Function
Echec:
    mErreurs = mErreurs & vbCrLf & "  " & nomFichier & " : " & Err.Description
    CreerPV = nomFichier & " (ÉCHEC)"
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

Private Sub RemplirPC(ws As Worksheet, d As Object, ByVal p As String)
    Ecrire ws, "E10", Lire(d, p & "MODELE")
    Ecrire ws, "E11", Lire(d, p & "SN")
    Ecrire ws, "E12", Lire(d, p & "NETBIOS")
    Ecrire ws, "F15", Lire(d, p & "STATION")
    Ecrire ws, "Q10", Lire(d, p & "ECRAN1")
    Ecrire ws, "Q12", Lire(d, p & "ECRAN2")
    Ecrire ws, "I49", UCase$(Lire(d, "NOM"))
    Ecrire ws, "P49", UCase$(Lire(d, "PRENOM"))
    Ecrire ws, "H51", TECHNICIEN
End Sub

' ---------------------------------------------------------------------
'  Charte iPhone (.xlsm + .pdf)
' ---------------------------------------------------------------------
Private Function CreerCharte(d As Object, ByVal base As String, ByVal dossier As String, _
                             ByVal nomFichier As String, ByVal dateInterv As Date) As String
    Dim wb As Workbook, ws As Worksheet, chemin As String, pdf As String

    If Len(Modele(base, "CHARTE")) = 0 Then
        mErreurs = mErreurs & vbCrLf & "  Charte vide introuvable dans : " & base
        CreerCharte = "(charte non créée)"
        Exit Function
    End If
    On Error GoTo Echec
    FermerSiOuvert Modele(base, "CHARTE")
    Set wb = Workbooks.Open(Filename:=Modele(base, "CHARTE"), UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set ws = wb.Sheets("Saisie")
    Ecrire ws, "C2", Lire(d, "IPI_IMEI")
    Ecrire ws, "C5", Lire(d, "IPI_SN")
    Ecrire ws, "C8", ModeleListe(wb, Lire(d, "IPI_MODELE"), "menu déroulant", "", "B2:B9")
    Ecrire ws, "C11", Operateur(Lire(d, "OPERATEUR"))
    Ecrire ws, "C14", UCase$(Lire(d, "NOM"))
    Ecrire ws, "C17", Prenom(Lire(d, "PRENOM"))
    Ecrire ws, "C20", Lire(d, "NNI")
    Ecrire ws, "C23", Lire(d, "LIGNE")
    ' Lieu, date et signature : manuscrits (réglage envoyé par le favori)
    AppliquerReglages wb, d, "Saisie"
    AppliquerReglages wb, d, "Prise en charge"
    wb.Application.Calculate

    chemin = CheminUnique(dossier, nomFichier, ".xlsm")
    wb.SaveAs Filename:=chemin, FileFormat:=52
    pdf = Left$(chemin, Len(chemin) - 5) & ".pdf"
    wb.Sheets("Prise en charge").ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdf, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    wb.Close SaveChanges:=False
    CreerCharte = Mid$(chemin, InStrRev(chemin, "\") + 1) & " + .pdf"
    Exit Function
Echec:
    mErreurs = mErreurs & vbCrLf & "  " & nomFichier & " : " & Err.Description
    CreerCharte = nomFichier & " (ÉCHEC)"
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

' Réglages envoyés par le favori, une ligne par cellule :
'   SET.<feuille>.<cellule>=valeur   -> écrit seulement si la cellule est vide
'   SETF.<feuille>.<cellule>=valeur  -> écrit toujours (remplace)
' Exemple : SET.Inst_PC.F17=Non   /   SETF.Prise en charge.B98=A : ... , le
Private Sub AppliquerReglages(wb As Workbook, d As Object, ByVal feuille As String)
    Dim k As Variant, cle As String, reste As String, p As Long, nomF As String, adr As String
    Dim force As Boolean, v As String
    For Each k In d.Keys
        cle = CStr(k)
        force = False
        If UCase$(Left$(cle, 5)) = "SETF." Then
            force = True: reste = Mid$(cle, 6)
        ElseIf UCase$(Left$(cle, 4)) = "SET." Then
            reste = Mid$(cle, 5)
        Else
            reste = ""
        End If
        If Len(reste) > 0 Then
            p = InStrRev(reste, ".")
            If p > 1 Then
                nomF = Left$(reste, p - 1)
                adr = Mid$(reste, p + 1)
                If StrComp(nomF, feuille, vbTextCompare) = 0 Then
                    v = CStr(d(k))
                    On Error Resume Next
                    If force Then
                        wb.Sheets(feuille).Range(adr).Value = v
                    ElseIf Len(Trim$(CStr(wb.Sheets(feuille).Range(adr).Value))) = 0 Then
                        Ecrire wb.Sheets(feuille), adr, v
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next k
End Sub

Private Function Operateur(ByVal v As String) As String
    Dim u As String
    u = UCase$(v)
    If InStr(u, "BOUYGUES") > 0 Then
        Operateur = "Bouygues"
    ElseIf InStr(u, "ORANGE") > 0 Then
        Operateur = "Orange"
    ElseIf InStr(u, "SFR") > 0 Then
        Operateur = "SFR"
    Else
        Operateur = v
    End If
End Function

Private Function Prenom(ByVal v As String) As String
    Prenom = StrConv(LCase$(Trim$(v)), vbProperCase)
End Function

' ---------------------------------------------------------------------
'  Journal (COSMOS\journal_pv.csv)
' ---------------------------------------------------------------------
Private Sub Journal(ByVal base As String, d As Object, ByVal dossier As String, ByVal fait As String)
    Dim n As Integer
    On Error Resume Next
    n = FreeFile
    Open base & "\journal_pv.csv" For Append As #n
    Print #n, Format(Now, "yyyy-mm-dd hh:nn") & ";" & Lire(d, "TICKET") & ";" & Lire(d, "DEM") & ";" & _
              Lire(d, "NOM") & " " & Lire(d, "PRENOM") & ";" & Lire(d, "ACTIONS") & ";" & dossier
    Close #n
End Sub

' ---------------------------------------------------------------------
'  Utilitaires
' ---------------------------------------------------------------------
Private Function LireUtf8(ByVal chemin As String) As String
    Dim st As Object
    On Error GoTo Secours
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile chemin
    LireUtf8 = st.ReadText
    st.Close
    Exit Function
Secours:
    Dim n As Integer, s As String
    n = FreeFile
    Open chemin For Binary As #n
    s = Space$(LOF(n))
    Get #n, , s
    Close #n
    LireUtf8 = s
End Function

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

' "JJ/MM/AAAA" -> Date (sinon aujourd'hui)
Private Function LireDate(ByVal s As String) As Date
    Dim p() As String
    LireDate = Date
    p = Split(Trim$(s), "/")
    If UBound(p) = 2 Then
        If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
            On Error Resume Next
            LireDate = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
        End If
    End If
End Function

Private Sub Ecrire(ws As Worksheet, ByVal adr As String, ByVal v As String)
    v = Trim$(v)
    If Len(v) = 0 Then Exit Sub
    On Error GoTo Erreur
    ' l'apostrophe force le texte (IMEI, n° de série, "014"...)
    ws.Range(adr).Value = "'" & v
    mNb = mNb + 1
    Exit Sub
Erreur:
    mErreurs = mErreurs & vbCrLf & "  " & ws.Name & "!" & adr & " non modifiable (" & v & ")"
End Sub

' Fait correspondre "APPLE iPhone 13 Pro" à une liste de modèles du classeur
Private Function ModeleListe(wb As Workbook, ByVal v As String, ByVal feuille As String, _
                             ByVal nomTable As String, ByVal plageSecours As String) As String
    Dim rng As Range, c As Range, nv As String, nc As String
    Dim best As String, bestLen As Long, court As String, nbCourt As Long

    v = NettoyerModele(v)
    ModeleListe = v
    If Len(v) = 0 Then Exit Function

    On Error Resume Next
    If Len(nomTable) > 0 Then Set rng = wb.Sheets(feuille).ListObjects(nomTable).DataBodyRange
    If rng Is Nothing Then Set rng = wb.Sheets(feuille).Range(plageSecours)
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
            If Left$(nc, Len(nv)) = nv Then
                court = CStr(c.Value): nbCourt = nbCourt + 1
            End If
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
