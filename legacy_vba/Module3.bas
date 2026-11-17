' --- MACRO PRINCIPALE ---
Sub CréerFeuilleCollection()
    Dim wsCollection As Worksheet, wsSommaire As Worksheet, wsArtiste As Worksheet
    Dim ligneCollection As Long, i As Long, colAlbum As Long
    Dim artiste As String, album As String, cellMoyenne As Range
    Dim dictCollection As Object, cleAlbum As String
    Dim valeurCD As String, valeurVinyle As String, derniereLigne As Long, j As Long
    Dim feuilleExiste As Boolean, dictAlbumsPossedes As Object, dictWishlist As Object
    Dim noteAlbum As Double, cValue As String, vValue As String
    
    On Error GoTo Fin_Erreur
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. TRIER TOUTES LES FEUILLES D'ARTISTES (Boucle renforcée)
    Call TrierToutesLesFeuillesArtistes

    ' 2. INITIALISATION
    Set wsSommaire = ThisWorkbook.Sheets("Sommaire")
    Set dictAlbumsPossedes = CreateObject("Scripting.Dictionary")
    Set dictWishlist = CreateObject("Scripting.Dictionary")
    Set dictCollection = CreateObject("Scripting.Dictionary")

    ' --- Gestion de la feuille Collection ---
    On Error Resume Next
    Set wsCollection = ThisWorkbook.Sheets("Collection")
    feuilleExiste = Not wsCollection Is Nothing
    On Error GoTo Fin_Erreur
    
    If feuilleExiste Then
        derniereLigne = wsCollection.Cells(wsCollection.Rows.Count, 1).End(xlUp).Row
        If derniereLigne > 1 Then
            For j = 2 To derniereLigne
                cleAlbum = wsCollection.Cells(j, 1).Value & " - " & wsCollection.Cells(j, 2).Value
                dictCollection(cleAlbum) = wsCollection.Cells(j, 3).Value & "|" & wsCollection.Cells(j, 4).Value
            Next j
        End If
        wsCollection.Cells.Clear
    Else
        Set wsCollection = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
        wsCollection.Name = "Collection"
    End If
    
    ' En-têtes Collection
    With wsCollection.Range("A1:D1")
        .Value = Array("Artiste", "Album", "CD", "Vinyle")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ligneCollection = 2
    
    ' 3. PARCOURS DE TOUTES LES FEUILLES POUR REMPLIR LA COLLECTION
    For Each wsArtiste In ThisWorkbook.Worksheets
        ' On exclut uniquement les pages de gestion
        If wsArtiste.Name <> "Sommaire" And wsArtiste.Name <> "Meilleurs Sons" And wsArtiste.Name <> "Collection" Then
            artiste = wsArtiste.Name
            
            For i = 0 To 39 ' Jusqu'à 40 albums
                colAlbum = 1 + (i * 5)
                Set cellMoyenne = wsArtiste.Cells(3, colAlbum + 3) ' Note de l'album
                
                If Not IsEmpty(wsArtiste.Cells(1, colAlbum)) Then ' Si un titre d'album existe
                    album = wsArtiste.Cells(1, colAlbum).Value
                    noteAlbum = 0
                    If IsNumeric(cellMoyenne.Value) Then noteAlbum = cellMoyenne.Value
                    
                    cleAlbum = artiste & " - " & album
                    
                    ' Récupération V/X existant
                    cValue = "X": vValue = "X"
                    If dictCollection.exists(cleAlbum) Then
                        Dim vals As Variant: vals = Split(dictCollection(cleAlbum), "|")
                        If vals(0) <> "" Then cValue = vals(0)
                        If vals(1) <> "" Then vValue = vals(1)
                    End If
                    
                    ' Logique Wishlist (W)
                    Dim wishCD As Boolean: wishCD = False
                    Dim wishVin As Boolean: wishVin = False
                    If noteAlbum > 8.5 And cValue <> "V" Then wishCD = True
                    If noteAlbum > 9 And vValue <> "V" Then wishVin = True
                    
                    ' Écriture Collection
                    wsCollection.Cells(ligneCollection, 1).Value = artiste
                    wsCollection.Cells(ligneCollection, 2).Value = album
                    wsCollection.Cells(ligneCollection, 3).Value = cValue
                    wsCollection.Cells(ligneCollection, 4).Value = vValue
                    ColoriserCellule wsCollection.Cells(ligneCollection, 3), cValue
                    ColoriserCellule wsCollection.Cells(ligneCollection, 4), vValue
                    
                    ' Stockage pour tableaux secondaires
                    If cValue = "V" Or vValue = "V" Then dictAlbumsPossedes(cleAlbum) = cValue & "|" & vValue
                    If wishCD Or wishVin Then dictWishlist(cleAlbum) = Array(wishCD, wishVin, cValue, vValue)
                    
                    ligneCollection = ligneCollection + 1
                End If
            Next i
        End If
    Next wsArtiste

    ' 4. GÉNÉRATION DES TABLEAUX À DROITE (F12 et K12)
    Call GenererTableauxSecondaires(wsCollection, dictAlbumsPossedes, dictWishlist)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Toutes les pages ont été triées et la collection est à jour !", vbInformation
    Exit Sub

Fin_Erreur:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur sur la feuille '" & artiste & "' : " & Err.Description
End Sub

' --- FONCTION DE TRI : PARCOURT TOUTES LES FEUILLES ---
Sub TrierToutesLesFeuillesArtistes()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' On trie TOUT ce qui n'est pas une page de menu
        If ws.Name <> "Sommaire" And ws.Name <> "Meilleurs Sons" And ws.Name <> "Collection" Then
            ' On ignore les erreurs sur une feuille précise pour continuer le tri sur les suivantes
            On Error Resume Next
            TrierUneFeuilleArtiste ws
            On Error GoTo 0
        End If
    Next ws
End Sub

' --- TRI D'UNE SEULE FEUILLE (A-E, F-J, K-O...) ---
Sub TrierUneFeuilleArtiste(ws As Worksheet)
    Dim i As Long, j As Long, nbAlbums As Long
    Dim col As Long, lastRow As Long
    Dim dates() As Variant, indices() As Long
    Dim tempVal As Variant, tempDate As Variant, tempIdx As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 10 Then lastRow = 50 ' Minimum de sécurité
    
    nbAlbums = 0
    ' On détecte combien d'albums sont présents
    For i = 0 To 39
        col = 1 + (i * 5)
        If ws.Cells(1, col).Value <> "" Then
            nbAlbums = nbAlbums + 1
            ReDim Preserve dates(1 To nbAlbums)
            ReDim Preserve indices(1 To nbAlbums)
            indices(nbAlbums) = col
            
            ' Lecture de l'année ou de la date en C1, H1, M1...
            tempVal = ws.Cells(1, col + 2).Value
            If IsNumeric(tempVal) And tempVal > 1900 And tempVal < 2100 Then
                dates(nbAlbums) = DateSerial(CInt(tempVal), 1, 1) ' Année seule
            ElseIf IsDate(tempVal) Then
                dates(nbAlbums) = CDate(tempVal) ' Date complète
            Else
                dates(nbAlbums) = DateSerial(9999, 12, 31) ' Pas de date = fin de liste
            End If
        End If
    Next i
    
    ' Si 0 ou 1 album, pas besoin de trier
    If nbAlbums < 2 Then Exit Sub
    
    ' Tri à bulles des colonnes par date
    For i = 1 To nbAlbums - 1
        For j = i + 1 To nbAlbums
            If dates(i) > dates(j) Then
                tempDate = dates(i): dates(i) = dates(j): dates(j) = tempDate
                tempIdx = indices(i): indices(i) = indices(j): indices(j) = tempIdx
            End If
        Next j
    Next i
    
    ' Création d'une zone temporaire pour réorganiser
    Dim wsTemp As Worksheet
    Set wsTemp = ThisWorkbook.Worksheets.Add
    
    For i = 1 To nbAlbums
        ' On copie le bloc de 5 colonnes (A-E devient le 1er, etc.)
        ws.Columns(indices(i)).Resize(, 5).Copy wsTemp.Columns(1 + ((i - 1) * 5))
    Next i
    
    ' On vide la feuille d'origine et on colle le résultat trié
    ws.Range("A1:ET" & lastRow + 20).Clear
    wsTemp.Range("A1:ET" & lastRow + 20).Copy ws.Range("A1")
    
    ' Supprimer la feuille temporaire
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True
End Sub

' --- GÉNÉRATION DES DEUX AUTRES TABLEAUX ---
Sub GenererTableauxSecondaires(ws As Worksheet, dPossedes As Object, dWish As Object)
    Dim lF12 As Long, lW As Long, cle As Variant, details As Variant, alb As Variant
    
    ' --- Tableau Possédés (F12) ---
    lF12 = 12
    ws.Range("F11:I11").Merge
    ws.Cells(11, 6).Value = "Albums Possédés (V)"
    ws.Range("F11:I11").Interior.Color = RGB(68, 114, 196)
    ws.Range("F11:I11").Font.Color = vbWhite
    ws.Range("F12:I12").Value = Array("Artiste", "Album", "CD", "Vinyle")
    ws.Range("F11:I12").Font.Bold = True
    
    For Each cle In dPossedes.Keys
        details = Split(dPossedes(cle), "|")
        alb = Split(cle, " - ")
        ws.Cells(lF12 + 1, 6).Value = alb(0)
        ws.Cells(lF12 + 1, 7).Value = alb(1)
        ws.Cells(lF12 + 1, 8).Value = IIf(details(0) = "V", "V", "")
        ws.Cells(lF12 + 1, 9).Value = IIf(details(1) = "V", "V", "")
        ColoriserCellule ws.Cells(lF12 + 1, 8), ws.Cells(lF12 + 1, 8).Value
        ColoriserCellule ws.Cells(lF12 + 1, 9), ws.Cells(lF12 + 1, 9).Value
        lF12 = lF12 + 1
    Next cle
    If lF12 > 12 Then ws.Range("F12:I" & lF12).Borders.LineStyle = xlContinuous

    ' --- Tableau Wishlist (K12) ---
    lW = 12
    ws.Range("K11:N11").Merge
    ws.Cells(11, 11).Value = "Wishlist (W)"
    ws.Range("K11:N11").Interior.Color = RGB(255, 140, 0)
    ws.Range("K11:N11").Font.Color = vbWhite
    ws.Range("K12:N12").Value = Array("Artiste", "Album", "CD (>8.5)", "Vinyle (>9)")
    ws.Range("K11:N12").Font.Bold = True
    
    For Each cle In dWish.Keys
        details = dWish(cle)
        alb = Split(cle, " - ")
        ws.Cells(lW + 1, 11).Value = alb(0)
        ws.Cells(lW + 1, 12).Value = alb(1)
        If details(0) And details(2) <> "V" Then ws.Cells(lW + 1, 13).Value = "W"
        If details(1) And details(3) <> "V" Then ws.Cells(lW + 1, 14).Value = "W"
        ColoriserCellule ws.Cells(lW + 1, 13), ws.Cells(lW + 1, 13).Value
        ColoriserCellule ws.Cells(lW + 1, 14), ws.Cells(lW + 1, 14).Value
        lW = lW + 1
    Next cle
    If lW > 12 Then ws.Range("K12:N" & lW).Borders.LineStyle = xlContinuous
    
    ws.Columns("A:N").AutoFit
End Sub

' --- COULEURS ---
Sub ColoriserCellule(cellule As Range, valeur As String)
    With cellule
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        Select Case UCase(Trim(valeur))
            Case "V": .Font.Color = RGB(0, 128, 0)
            Case "X": .Font.Color = RGB(255, 0, 0)
            Case "W": .Font.Color = RGB(255, 140, 0)
            Case Else: .Font.Bold = False: .Font.Color = vbBlack
        End Select
    End With
End Sub
