Sub CréerMettreAJourSommaire()

    Dim wsSommaire As Worksheet
    Dim wsArtiste As Worksheet
    Dim dictAlbums As Object
    Dim dictArtistes As Object
    Dim dictStyles As Object
    
    ' Nouveaux Dictionnaires pour le RAP (artistes)
    Dim dictArtistesRAPFR As Object
    Dim dictArtistesRAPUS As Object
    
    ' Nouveau Dictionnaire pour les Meilleurs Sons par Style
    Dim dictMeilleursSonsParStyle As Object
    
    Dim arrAlbums() As Variant
    Dim arrArtistes() As Variant
    Dim arrStyles() As Variant
    Dim arrArtistesRAPFR() As Variant
    Dim arrArtistesRAPUS() As Variant
    
    Dim i As Long, j As Long, k As Long
    Dim temp As Variant
    Dim colArtistes As Long
    Dim colAlbums As Long
    Dim colStyles As Long
    Dim ligneClassement As Long
    Dim artiste As String
    Dim album As String
    Dim styleMusical As Variant
    Dim moyenneAlbum As Variant
    Dim cellMoyenne As Range
    Dim albumDetails As Variant
    Dim Key As Variant
    Dim lignePires As Long
    Dim sommeNotesArtiste As Double
    Dim compteurAlbumsArtiste As Long
    Dim colAlbum As Long
    Dim moyenneArtiste As Double
    Dim compteurGlobal As Long
    Dim lastRowSommaire As Long
    Dim titreChanson As String
    Dim lastRowArtiste As Long
    Dim ligneSonsEtoile As Long
    Dim styleAlbum As String
    
    ' Variables pour les sons RAP
    Dim collectionSonsRAP As New Collection

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- FORMATTAGE DES NOTES AVEC DEUX DÉCIMALES ---
    For Each wsArtiste In ThisWorkbook.Worksheets
        If wsArtiste.Name <> "Sommaire" And wsArtiste.Name <> "Meilleurs Sons" And wsArtiste.Visible = xlSheetVisible Then
            For i = 0 To 39
                colAlbum = 1 + (i * 5)
                Set cellMoyenne = wsArtiste.Cells(3, colAlbum + 3)
                If Not IsEmpty(cellMoyenne) And IsNumeric(cellMoyenne.Value) Then
                    cellMoyenne.NumberFormat = "0.00"
                End If
            Next i
        End If
    Next wsArtiste

    ' --- Création ou vidage de la feuille Sommaire ---
    On Error Resume Next
    Set wsSommaire = ThisWorkbook.Sheets("Sommaire")
    If wsSommaire Is Nothing Then
        Set wsSommaire = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        wsSommaire.Name = "Sommaire"
    Else
        wsSommaire.Cells.Clear
    End If
    On Error GoTo 0
    
    Set dictAlbums = CreateObject("Scripting.Dictionary")
    Set dictArtistes = CreateObject("Scripting.Dictionary")
    Set dictStyles = CreateObject("Scripting.Dictionary")
    Set dictArtistesRAPFR = CreateObject("Scripting.Dictionary")
    Set dictArtistesRAPUS = CreateObject("Scripting.Dictionary")
    Set dictMeilleursSonsParStyle = CreateObject("Scripting.Dictionary")
    
    ' --- 1. COLLECTE DES DONNÉES, COMPTAGE ET SONS ÉTOILÉS PAR STYLE ---
    For Each wsArtiste In ThisWorkbook.Worksheets
        If wsArtiste.Name <> "Sommaire" And wsArtiste.Visible = xlSheetVisible Then
            artiste = wsArtiste.Name
            sommeNotesArtiste = 0
            compteurAlbumsArtiste = 0
            
            Dim dictStylesArtiste As Object
            Set dictStylesArtiste = CreateObject("Scripting.Dictionary")
            
            For i = 0 To 39
                colAlbum = 1 + (i * 5)
                Set cellMoyenne = wsArtiste.Cells(3, colAlbum + 3)
                
                If Not IsEmpty(cellMoyenne) And IsNumeric(cellMoyenne.Value) Then
                    album = wsArtiste.Cells(1, colAlbum).Value
                    styleMusical = wsArtiste.Cells(1, colAlbum + 1).Value ' Style de l'album
                    moyenneAlbum = cellMoyenne.Value
                    
                    ' --- COLLECTE DES ALBUMS ET STYLES ---
                    If Not dictAlbums.exists(artiste & " - " & album) Then
                        dictAlbums.Add artiste & " - " & album, moyenneAlbum
                    End If
                    
                    If styleMusical <> "" Then
                        If Not dictStyles.exists(styleMusical) Then
                            Set dictStyles.Item(styleMusical) = CreateObject("Scripting.Dictionary")
                        End If
                        If Not dictStyles(styleMusical).exists(artiste & " - " & album) Then
                            dictStyles(styleMusical).Add artiste & " - " & album, moyenneAlbum
                        End If
                        
                        If Not dictStylesArtiste.exists(styleMusical) Then
                             dictStylesArtiste.Add styleMusical, True
                        End If
                    End If
                    
                    sommeNotesArtiste = sommeNotesArtiste + moyenneAlbum
                    compteurAlbumsArtiste = compteurAlbumsArtiste + 1
                End If
                
                ' --- COLLECTE DES SONS ÉTOILÉS PAR STYLE MUSICAL ---
                album = wsArtiste.Cells(1, colAlbum).Value
                styleAlbum = wsArtiste.Cells(1, colAlbum + 1).Value
                
                If Not IsEmpty(album) Then
                    lastRowArtiste = wsArtiste.Cells(Rows.Count, colAlbum).End(xlUp).Row
                    
                    For j = 4 To lastRowArtiste
                        titreChanson = CStr(wsArtiste.Cells(j, colAlbum).Value)
                        
                        If Len(titreChanson) > 0 And Right(titreChanson, 1) = "*" Then
                            
                            If styleAlbum <> "" Then
                                If Not dictMeilleursSonsParStyle.exists(styleAlbum) Then
                                    ' Créer une collection pour stocker les chansons de ce style
                                    Set dictMeilleursSonsParStyle.Item(styleAlbum) = New Collection
                                End If
                                ' Ajouter la chanson au format "Artiste - Titre"
                                dictMeilleursSonsParStyle(styleAlbum).Add Array(artiste, Left(titreChanson, Len(titreChanson) - 1))
                            End If
                        End If
                    Next j
                End If

            Next i
            
            ' --- CALCUL MOYENNE ARTISTE ---
            If compteurAlbumsArtiste > 0 Then
                moyenneArtiste = sommeNotesArtiste / compteurAlbumsArtiste
                If Not dictArtistes.exists(artiste) Then
                    dictArtistes.Add artiste, moyenneArtiste
                End If
                
                ' --- LOGIQUE RAP FR/US POUR LE CLASSEMENT ARTISTE ---
                If dictStylesArtiste.exists("RAP FR") Then
                    If Not dictArtistesRAPFR.exists(artiste) Then
                        dictArtistesRAPFR.Add artiste, moyenneArtiste
                    End If
                End If
                If dictStylesArtiste.exists("RAP US") Then
                    If Not dictArtistesRAPUS.exists(artiste) Then
                        dictArtistesRAPUS.Add artiste, moyenneArtiste
                    End If
                End If
            End If
        End If
    Next wsArtiste
    
    ' --- 2. MISE EN PAGE DU SOMMAIRE (inchangée) ---
    ' ... (Le code de la section 2 pour les statistiques/mise en page générale est conservé)
    With wsSommaire
        .Cells(3, 1).Value = "Albums notés :"
        .Cells(3, 2).Value = dictAlbums.Count
        
        .Cells(4, 1).Value = "Artistes notés :"
        .Cells(4, 2).Value = dictArtistes.Count
        
        .Cells(5, 1).Value = "Albums par style :"
        
        compteurGlobal = 6
        For Each Key In dictStyles.Keys
            .Cells(compteurGlobal, 1).Value = Key & ":"
            .Cells(compteurGlobal, 2).Value = dictStyles(Key).Count
            compteurGlobal = compteurGlobal + 1
        Next Key
        
        .Cells(compteurGlobal + 1, 1).Value = "Liste des artistes"
        .Cells(compteurGlobal + 1, 1).Font.Bold = True
        .Cells(compteurGlobal + 1, 1).Font.Size = 12
        Dim ligneArtiste As Long
        ligneArtiste = compteurGlobal + 2
        For Each wsArtiste In ThisWorkbook.Worksheets
            If wsArtiste.Name <> "Sommaire" And wsArtiste.Visible = xlSheetVisible Then
                .Cells(ligneArtiste, 1).Value = wsArtiste.Name
                .Hyperlinks.Add Anchor:=.Cells(ligneArtiste, 1), _
                                 Address:="", _
                                 SubAddress:="'" & wsArtiste.Name & "'!A1", _
                                 TextToDisplay:=wsArtiste.Name
                ligneArtiste = ligneArtiste + 1
            End If
        Next wsArtiste
        
        ' Ajout d'une colonne de classement pour les tops (Artistes Globaux)
        colArtistes = 4
        .Cells(4, colArtistes).Value = "Top 20 des meilleurs artistes"
        .Cells(5, colArtistes).Value = "Place" ' Nouvelle colonne
        .Cells(5, colArtistes + 1).Value = "Note"
        .Cells(5, colArtistes + 2).Value = "Artiste"
        With .Range(.Cells(4, colArtistes), .Cells(5, colArtistes + 2))
            .Font.Bold = True: .Interior.Color = RGB(0, 112, 192)
            .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter
        End With
        .Cells(4, colArtistes).MergeCells = True
        
        colAlbums = 8
        .Cells(4, colAlbums).Value = "Top 20 des meilleurs albums"
        .Cells(5, colAlbums).Value = "Place" ' Nouvelle colonne
        .Cells(5, colAlbums + 1).Value = "Note"
        .Cells(5, colAlbums + 2).Value = "Artiste"
        .Cells(5, colAlbums + 3).Value = "Album"
        With .Range(.Cells(4, colAlbums), .Cells(5, colAlbums + 3))
            .Font.Bold = True: .Interior.Color = RGB(112, 173, 71)
            .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter
        End With
        .Cells(4, colAlbums).MergeCells = True
        
    End With
    
    ' --- 3. PEUPLEMENT DES TOPS (Artistes Globaux et Albums) ---
    ' ... (Le code de tri et d'affichage des tops Artistes et Albums globaux est conservé)
    
    ' --- Tops Artistes Global (décroissant) ---
    ReDim arrArtistes(1 To dictArtistes.Count, 1 To 2)
    i = 1
    For Each Key In dictArtistes.Keys
        arrArtistes(i, 1) = dictArtistes.Item(Key)
        arrArtistes(i, 2) = Key
        i = i + 1
    Next Key
    
    For i = LBound(arrArtistes, 1) To UBound(arrArtistes, 1) - 1
        For j = i + 1 To UBound(arrArtistes, 1)
            If arrArtistes(i, 1) < arrArtistes(j, 1) Then
                temp = arrArtistes(i, 1): arrArtistes(i, 1) = arrArtistes(j, 1): arrArtistes(j, 1) = temp
                temp = arrArtistes(i, 2): arrArtistes(i, 2) = arrArtistes(j, 2): arrArtistes(j, 2) = temp
            End If
        Next j
    Next i
    
    ligneClassement = 6: colArtistes = 4
    For i = 1 To WorksheetFunction.Min(20, UBound(arrArtistes, 1))
        With wsSommaire
            .Cells(ligneClassement, colArtistes).Value = i & "e"
            .Cells(ligneClassement, colArtistes + 1).Value = arrArtistes(i, 1)
            .Cells(ligneClassement, colArtistes + 2).Value = arrArtistes(i, 2)
            .Cells(ligneClassement, colArtistes + 1).NumberFormat = "0.00"
            
            If arrArtistes(i, 1) >= 9 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(0, 153, 0)
            ElseIf arrArtistes(i, 1) >= 8 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(146, 208, 80)
            ElseIf arrArtistes(i, 1) >= 7 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(173, 226, 85)
            ElseIf arrArtistes(i, 1) >= 5.5 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(255, 255, 0)
            ElseIf arrArtistes(i, 1) >= 4 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(255, 192, 0)
            Else
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(255, 0, 0)
            End If
        End With
        ligneClassement = ligneClassement + 1
    Next i
    
    ' --- Top 20 des pires artistes Global (croissant) ---
    For i = LBound(arrArtistes, 1) To UBound(arrArtistes, 1) - 1
        For j = i + 1 To UBound(arrArtistes, 1)
            If arrArtistes(i, 1) > arrArtistes(j, 1) Then
                temp = arrArtistes(i, 1): arrArtistes(i, 1) = arrArtistes(j, 1): arrArtistes(j, 1) = temp
                temp = arrArtistes(i, 2): arrArtistes(i, 2) = arrArtistes(j, 2): arrArtistes(j, 2) = temp
            End If
        Next j
    Next i
    
    ligneClassement = 29: colArtistes = 4
    With wsSommaire
        .Cells(ligneClassement, colArtistes).Value = "Top 20 des pires artistes"
        .Cells(ligneClassement + 1, colArtistes).Value = "Place"
        .Cells(ligneClassement + 1, colArtistes + 1).Value = "Note"
        .Cells(ligneClassement + 1, colArtistes + 2).Value = "Artiste"
        With .Range(.Cells(ligneClassement, colArtistes), .Cells(ligneClassement + 1, colArtistes + 2))
            .Font.Bold = True: .Interior.Color = RGB(255, 153, 0)
            .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter
        End With
        .Cells(ligneClassement, colArtistes).MergeCells = True
    End With

    ligneClassement = 31: colArtistes = 4
    For i = 1 To WorksheetFunction.Min(20, UBound(arrArtistes, 1))
        With wsSommaire
            .Cells(ligneClassement, colArtistes).Value = i & "e"
            .Cells(ligneClassement, colArtistes + 1).Value = arrArtistes(i, 1)
            .Cells(ligneClassement, colArtistes + 2).Value = arrArtistes(i, 2)
            .Cells(ligneClassement, colArtistes + 1).NumberFormat = "0.00"
            
            If arrArtistes(i, 1) >= 9 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(0, 153, 0)
            ElseIf arrArtistes(i, 1) >= 8 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(146, 208, 80)
            ElseIf arrArtistes(i, 1) >= 7 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(173, 226, 85)
            ElseIf arrArtistes(i, 1) >= 5.5 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(255, 255, 0)
            ElseIf arrArtistes(i, 1) >= 4 Then
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(255, 192, 0)
            Else
                .Cells(ligneClassement, colArtistes + 1).Interior.Color = RGB(255, 0, 0)
            End If
        End With
        ligneClassement = ligneClassement + 1
    Next i
    
    ' --- Tops Albums Global (décroissant) ---
    ReDim arrAlbums(1 To dictAlbums.Count, 1 To 2)
    i = 1
    For Each Key In dictAlbums.Keys
        arrAlbums(i, 1) = dictAlbums.Item(Key)
        arrAlbums(i, 2) = Key
        i = i + 1
    Next Key
    
    For i = LBound(arrAlbums, 1) To UBound(arrAlbums, 1) - 1
        For j = i + 1 To UBound(arrAlbums, 1)
            If arrAlbums(i, 1) < arrAlbums(j, 1) Then
                temp = arrAlbums(i, 1): arrAlbums(i, 1) = arrAlbums(j, 1): arrAlbums(j, 1) = temp
                temp = arrAlbums(i, 2): arrAlbums(i, 2) = arrAlbums(j, 2): arrAlbums(j, 2) = temp
            End If
        Next j
    Next i
    
    ligneClassement = 6: colAlbums = 8
    For i = 1 To WorksheetFunction.Min(20, UBound(arrAlbums, 1))
        albumDetails = Split(arrAlbums(i, 2), " - ")
        With wsSommaire
            .Cells(ligneClassement, colAlbums).Value = i & "e"
            .Cells(ligneClassement, colAlbums + 1).Value = arrAlbums(i, 1)
            .Cells(ligneClassement, colAlbums + 2).Value = albumDetails(0)
            .Cells(ligneClassement, colAlbums + 3).Value = albumDetails(1)
            .Cells(ligneClassement, colAlbums + 1).NumberFormat = "0.00"
            
            .Hyperlinks.Add Anchor:=.Cells(ligneClassement, colAlbums + 3), _
                             Address:="", _
                             SubAddress:="'" & albumDetails(0) & "'!A1", _
                             TextToDisplay:=albumDetails(1)
            
            If arrAlbums(i, 1) >= 9 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(0, 153, 0)
            ElseIf arrAlbums(i, 1) >= 8 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(146, 208, 80)
            ElseIf arrAlbums(i, 1) >= 7 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(173, 226, 85)
            ElseIf arrAlbums(i, 1) >= 5.5 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(255, 255, 0)
            ElseIf arrAlbums(i, 1) >= 4 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(255, 192, 0)
            Else
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(255, 0, 0)
            End If
        End With
        ligneClassement = ligneClassement + 1
    Next i
    
    ' --- Top 20 des pires albums Global (croissant) ---
    For i = LBound(arrAlbums, 1) To UBound(arrAlbums, 1) - 1
        For j = i + 1 To UBound(arrAlbums, 1)
            If arrAlbums(i, 1) > arrAlbums(j, 1) Then
                temp = arrAlbums(i, 1): arrAlbums(i, 1) = arrAlbums(j, 1): arrAlbums(j, 1) = temp
                temp = arrAlbums(i, 2): arrAlbums(i, 2) = arrAlbums(j, 2): arrAlbums(j, 2) = temp
            End If
        Next j
    Next i
    
    ligneClassement = 29: colAlbums = 8
    With wsSommaire
        .Cells(ligneClassement, colAlbums).Value = "Top 20 des pires albums"
        .Cells(ligneClassement + 1, colAlbums).Value = "Place"
        .Cells(ligneClassement + 1, colAlbums + 1).Value = "Note"
        .Cells(ligneClassement + 1, colAlbums + 2).Value = "Artiste"
        .Cells(ligneClassement + 1, colAlbums + 3).Value = "Album"
        
        With .Range(.Cells(ligneClassement, colAlbums), .Cells(ligneClassement + 1, colAlbums + 3))
            .Font.Bold = True: .Interior.Color = RGB(192, 0, 0)
            .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter
        End With
        .Cells(ligneClassement, colAlbums).MergeCells = True
    End With

    ligneClassement = 31: colAlbums = 8
    For i = 1 To WorksheetFunction.Min(20, UBound(arrAlbums, 1))
        albumDetails = Split(arrAlbums(i, 2), " - ")
        With wsSommaire
            .Cells(ligneClassement, colAlbums).Value = i & "e"
            .Cells(ligneClassement, colAlbums + 1).Value = arrAlbums(i, 1)
            .Cells(ligneClassement, colAlbums + 2).Value = albumDetails(0)
            .Cells(ligneClassement, colAlbums + 3).Value = albumDetails(1)
            .Cells(ligneClassement, colAlbums + 1).NumberFormat = "0.00"
            
            .Hyperlinks.Add Anchor:=.Cells(ligneClassement, colAlbums + 3), _
                             Address:="", _
                             SubAddress:="'" & albumDetails(0) & "'!A1", _
                             TextToDisplay:=albumDetails(1)
            
            If arrAlbums(i, 1) >= 9 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(0, 153, 0)
            ElseIf arrAlbums(i, 1) >= 8 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(146, 208, 80)
            ElseIf arrAlbums(i, 1) >= 7 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(173, 226, 85)
            ElseIf arrAlbums(i, 1) >= 5.5 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(255, 255, 0)
            ElseIf arrAlbums(i, 1) >= 4 Then
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(255, 192, 0)
            Else
                .Cells(ligneClassement, colAlbums + 1).Interior.Color = RGB(255, 0, 0)
            End If
        End With
        ligneClassement = ligneClassement + 1
    Next i

    ' --- NOUVEAU : CLASSEMENTS ARTISTES RAP FR/US ---
    colArtistes = 12 ' Nouvelle colonne de départ (après les tops albums)
    
    ' --- Top Artistes RAP FR ---
    If dictArtistesRAPFR.Count > 0 Then
        ReDim arrArtistesRAPFR(1 To dictArtistesRAPFR.Count, 1 To 2)
        i = 1
        For Each Key In dictArtistesRAPFR.Keys
            arrArtistesRAPFR(i, 1) = dictArtistesRAPFR.Item(Key)
            arrArtistesRAPFR(i, 2) = Key
            i = i + 1
        Next Key
        
        ' Tri décroissant
        For i = LBound(arrArtistesRAPFR, 1) To UBound(arrArtistesRAPFR, 1) - 1
            For j = i + 1 To UBound(arrArtistesRAPFR, 1)
                If arrArtistesRAPFR(i, 1) < arrArtistesRAPFR(j, 1) Then
                    temp = arrArtistesRAPFR(i, 1): arrArtistesRAPFR(i, 1) = arrArtistesRAPFR(j, 1): arrArtistesRAPFR(j, 1) = temp
                    temp = arrArtistesRAPFR(i, 2): arrArtistesRAPFR(i, 2) = arrArtistesRAPFR(j, 2): arrArtistesRAPFR(j, 2) = temp
                End If
            Next j
        Next i
        
        ligneClassement = 4: colArtistes = colArtistes
        With wsSommaire
            .Cells(ligneClassement, colArtistes).Value = "Top Artistes RAP FR"
            .Cells(ligneClassement + 1, colArtistes).Value = "Place"
            .Cells(ligneClassement + 1, colArtistes + 1).Value = "Note"
            .Cells(ligneClassement + 1, colArtistes + 2).Value = "Artiste"
            With .Range(.Cells(ligneClassement, colArtistes), .Cells(ligneClassement + 1, colArtistes + 2))
                .Font.Bold = True: .Interior.Color = RGB(132, 132, 132) ' Gris foncé
                .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter
            End With
            .Cells(ligneClassement, colArtistes).MergeCells = True
        End With
        
        ligneClassement = 6
        For i = 1 To WorksheetFunction.Min(20, UBound(arrArtistesRAPFR, 1))
            With wsSommaire
                .Cells(ligneClassement, colArtistes).Value = i & "e"
                .Cells(ligneClassement, colArtistes + 1).Value = arrArtistesRAPFR(i, 1)
                .Cells(ligneClassement, colArtistes + 2).Value = arrArtistesRAPFR(i, 2)
                .Cells(ligneClassement, colArtistes + 1).NumberFormat = "0.00"
            End With
            ligneClassement = ligneClassement + 1
        Next i
        
        colArtistes = colArtistes + 4 ' Décalage pour la prochaine colonne de classement
    End If
    
    ' --- Top Artistes RAP US ---
    If dictArtistesRAPUS.Count > 0 Then
        ReDim arrArtistesRAPUS(1 To dictArtistesRAPUS.Count, 1 To 2)
        i = 1
        For Each Key In dictArtistesRAPUS.Keys
            arrArtistesRAPUS(i, 1) = dictArtistesRAPUS.Item(Key)
            arrArtistesRAPUS(i, 2) = Key
            i = i + 1
        Next Key
        
        ' Tri décroissant
        For i = LBound(arrArtistesRAPUS, 1) To UBound(arrArtistesRAPUS, 1) - 1
            For j = i + 1 To UBound(arrArtistesRAPUS, 1)
                If arrArtistesRAPUS(i, 1) < arrArtistesRAPUS(j, 1) Then
                    temp = arrArtistesRAPUS(i, 1): arrArtistesRAPUS(i, 1) = arrArtistesRAPUS(j, 1): arrArtistesRAPUS(j, 1) = temp
                    temp = arrArtistesRAPUS(i, 2): arrArtistesRAPUS(i, 2) = arrArtistesRAPUS(j, 2): arrArtistesRAPUS(j, 2) = temp
                End If
            Next j
        Next i
        
        ligneClassement = 4:
        With wsSommaire
            .Cells(ligneClassement, colArtistes).Value = "Top Artistes RAP US"
            .Cells(ligneClassement + 1, colArtistes).Value = "Place"
            .Cells(ligneClassement + 1, colArtistes + 1).Value = "Note"
            .Cells(ligneClassement + 1, colArtistes + 2).Value = "Artiste"
            With .Range(.Cells(ligneClassement, colArtistes), .Cells(ligneClassement + 1, colArtistes + 2))
                .Font.Bold = True: .Interior.Color = RGB(178, 34, 34) ' Rouge Brique
                .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter
            End With
            .Cells(ligneClassement, colArtistes).MergeCells = True
        End With
        
        ligneClassement = 6
        For i = 1 To WorksheetFunction.Min(20, UBound(arrArtistesRAPUS, 1))
            With wsSommaire
                .Cells(ligneClassement, colArtistes).Value = i & "e"
                .Cells(ligneClassement, colArtistes + 1).Value = arrArtistesRAPUS(i, 1)
                .Cells(ligneClassement, colArtistes + 2).Value = arrArtistesRAPUS(i, 2)
                .Cells(ligneClassement, colArtistes + 1).NumberFormat = "0.00"
            End With
            ligneClassement = ligneClassement + 1
        Next i
        
        colArtistes = colArtistes + 4 ' Décalage pour la prochaine colonne de classement
    End If
    
    ' --- 4. TOPS ALBUMS PAR STYLE MUSICAL (inchangé) ---
    colStyles = colArtistes ' Commence après les classements RAP
    
    For Each styleMusical In dictStyles.Keys
        ligneClassement = 4
        
        ReDim arrStyles(1 To dictStyles(styleMusical).Count, 1 To 2)
        i = 1
        For Each Key In dictStyles(styleMusical).Keys
            arrStyles(i, 1) = dictStyles(styleMusical).Item(Key)
            arrStyles(i, 2) = Key
            i = i + 1
        Next Key
        
        For i = LBound(arrStyles, 1) To UBound(arrStyles, 1) - 1
            For j = i + 1 To UBound(arrStyles, 1)
                If arrStyles(i, 1) < arrStyles(j, 1) Then
                    temp = arrStyles(i, 1): arrStyles(i, 1) = arrStyles(j, 1): arrStyles(j, 1) = temp
                    temp = arrStyles(i, 2): arrStyles(i, 2) = arrStyles(j, 2): arrStyles(j, 2) = temp
                End If
            Next j
        Next i
        
        With wsSommaire
            .Cells(ligneClassement, colStyles).Value = "Top 10 " & styleMusical
            .Cells(ligneClassement + 1, colStyles).Value = "Place"
            .Cells(ligneClassement + 1, colStyles + 1).Value = "Note"
            .Cells(ligneClassement + 1, colStyles + 2).Value = "Artiste"
            .Cells(ligneClassement + 1, colStyles + 3).Value = "Album"
            
            With .Range(.Cells(ligneClassement, colStyles), .Cells(ligneClassement + 1, colStyles + 3))
                .Font.Bold = True: .Interior.Color = RGB(155, 194, 230)
                .Font.Color = RGB(0, 0, 0): .HorizontalAlignment = xlCenter
            End With
            .Cells(ligneClassement, colStyles).MergeCells = True
            
            ligneClassement = 6
            For i = 1 To WorksheetFunction.Min(10, UBound(arrStyles, 1))
                albumDetails = Split(arrStyles(i, 2), " - ")
                .Cells(ligneClassement, colStyles).Value = i & "e"
                .Cells(ligneClassement, colStyles + 1).Value = arrStyles(i, 1)
                .Cells(ligneClassement, colStyles + 2).Value = albumDetails(0)
                .Cells(ligneClassement, colStyles + 3).Value = albumDetails(1)
                .Cells(ligneClassement, colStyles + 1).NumberFormat = "0.00"
                
                .Hyperlinks.Add Anchor:=.Cells(ligneClassement, colStyles + 3), _
                                 Address:="", _
                                 SubAddress:="'" & albumDetails(0) & "'!A1", _
                                 TextToDisplay:=albumDetails(1)
                
                If arrStyles(i, 1) >= 9 Then
                    .Cells(ligneClassement, colStyles + 1).Interior.Color = RGB(0, 153, 0)
                ElseIf arrStyles(i, 1) >= 8 Then
                    .Cells(ligneClassement, colStyles + 1).Interior.Color = RGB(146, 208, 80)
                ElseIf arrStyles(i, 1) >= 7 Then
                    .Cells(ligneClassement, colStyles + 1).Interior.Color = RGB(173, 226, 85)
                ElseIf arrStyles(i, 1) >= 5.5 Then
                    .Cells(ligneClassement, colStyles + 1).Interior.Color = RGB(255, 255, 0)
                ElseIf arrStyles(i, 1) >= 4 Then
                    .Cells(ligneClassement, colStyles + 1).Interior.Color = RGB(255, 192, 0)
                Else
                    .Cells(ligneClassement, colStyles + 1).Interior.Color = RGB(255, 0, 0)
                End If
                
                ligneClassement = ligneClassement + 1
            Next i
        End With
        
        For i = LBound(arrStyles, 1) To UBound(arrStyles, 1) - 1
            For j = i + 1 To UBound(arrStyles, 1)
                If arrStyles(i, 1) > arrStyles(j, 1) Then
                    temp = arrStyles(i, 1): arrStyles(i, 1) = arrStyles(j, 1): arrStyles(j, 1) = temp
                    temp = arrStyles(i, 2): arrStyles(i, 2) = arrStyles(j, 2): arrStyles(j, 2) = temp
                End If
            Next j
        Next i
        
        With wsSommaire
            .Cells(ligneClassement + 2, colStyles).Value = "Top 10 pires " & styleMusical
            .Cells(ligneClassement + 3, colStyles).Value = "Place"
            .Cells(ligneClassement + 3, colStyles + 1).Value = "Note"
            .Cells(ligneClassement + 3, colStyles + 2).Value = "Artiste"
            .Cells(ligneClassement + 3, colStyles + 3).Value = "Album"
            
            With .Range(.Cells(ligneClassement + 2, colStyles), .Cells(ligneClassement + 3, colStyles + 3))
                .Font.Bold = True: .Interior.Color = RGB(230, 155, 194)
                .Font.Color = RGB(0, 0, 0): .HorizontalAlignment = xlCenter
            End With
            .Cells(ligneClassement + 2, colStyles).MergeCells = True
            
            lignePires = ligneClassement + 4
            For i = 1 To WorksheetFunction.Min(10, UBound(arrStyles, 1))
                albumDetails = Split(arrStyles(i, 2), " - ")
                .Cells(lignePires, colStyles).Value = i & "e"
                .Cells(lignePires, colStyles + 1).Value = arrStyles(i, 1)
                .Cells(lignePires, colStyles + 2).Value = albumDetails(0)
                .Cells(lignePires, colStyles + 3).Value = albumDetails(1)
                .Cells(lignePires, colStyles + 1).NumberFormat = "0.00"
                
                .Hyperlinks.Add Anchor:=.Cells(lignePires, colStyles + 3), _
                                 Address:="", _
                                 SubAddress:="'" & albumDetails(0) & "'!A1", _
                                 TextToDisplay:=albumDetails(1)
                
                If arrStyles(i, 1) >= 9 Then
                    .Cells(lignePires, colStyles + 1).Interior.Color = RGB(0, 153, 0)
                ElseIf arrStyles(i, 1) >= 8 Then
                    .Cells(lignePires, colStyles + 1).Interior.Color = RGB(146, 208, 80)
                ElseIf arrStyles(i, 1) >= 7 Then
                    .Cells(lignePires, colStyles + 1).Interior.Color = RGB(173, 226, 85)
                ElseIf arrStyles(i, 1) >= 5.5 Then
                    .Cells(lignePires, colStyles + 1).Interior.Color = RGB(255, 255, 0)
                ElseIf arrStyles(i, 1) >= 4 Then
                    .Cells(lignePires, colStyles + 1).Interior.Color = RGB(255, 192, 0)
                Else
                    .Cells(lignePires, colStyles + 1).Interior.Color = RGB(255, 0, 0)
                End If
                
                lignePires = lignePires + 1
            Next i
        End With
        
        colStyles = colStyles + 5
    Next styleMusical
    
    ' --- 5. NOUVEAU : AFFICHAGE DES MEILLEURS SONS PAR STYLE MUSICAL ---
    
    ' Trouver la dernière ligne utilisée pour placer la nouvelle section
    lastRowSommaire = wsSommaire.Cells(Rows.Count, "A").End(xlUp).Row + 5
    
    colArtistes = 1 ' Colonne de départ pour la nouvelle section

    For Each Key In dictMeilleursSonsParStyle.Keys
        
        styleMusical = Key
        Dim sonsCollection As Collection
        Set sonsCollection = dictMeilleursSonsParStyle(styleMusical)
        
        If sonsCollection.Count > 0 Then
            With wsSommaire
                ' Titre de la section
                .Cells(lastRowSommaire, colArtistes).Value = "Meilleurs Sons : " & styleMusical
                .Cells(lastRowSommaire, colArtistes).Font.Bold = True
                .Cells(lastRowSommaire, colArtistes).Font.Size = 12
                
                ' En-têtes
                .Cells(lastRowSommaire + 1, colArtistes).Value = "Artiste"
                .Cells(lastRowSommaire + 1, colArtistes + 1).Value = "Titre de la chanson"
                .Range(.Cells(lastRowSommaire + 1, colArtistes), .Cells(lastRowSommaire + 1, colArtistes + 1)).Font.Bold = True
                .Range(.Cells(lastRowSommaire + 1, colArtistes), .Cells(lastRowSommaire + 1, colArtistes + 1)).Interior.Color = RGB(220, 220, 220)
                
                ligneSonsEtoile = lastRowSommaire + 2
                
                ' Affichage des chansons
                For Each temp In sonsCollection
                    .Cells(ligneSonsEtoile, colArtistes).Value = temp(0) ' Artiste
                    .Cells(ligneSonsEtoile, colArtistes + 1).Value = temp(1) ' Titre
                    ligneSonsEtoile = ligneSonsEtoile + 1
                Next temp
                
                lastRowSommaire = ligneSonsEtoile + 2 ' Préparer la prochaine position
            End With
        End If
        
        ' Déplacer la colonne de départ pour éviter que les listes ne se chevauchent si on veut les mettre côte à côte
        ' colArtistes = colArtistes + 3 ' Utilise cette ligne si tu veux aligner les listes en colonnes
        
    Next Key
    
    ' --- Mise en forme finale ---
    With wsSommaire
        .Columns.AutoFit
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Le sommaire a été mis à jour avec succès !", vbInformation, "Sommaire"

Fin:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

Fin_Erreur:
    MsgBox "Une erreur s'est produite : " & Err.Description, vbCritical, "Erreur"
    Resume Fin

End Sub

