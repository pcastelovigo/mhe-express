' © Pablo Manuel Castelo Vigo 2021 -- pablocv@gmx.es
' Licenciado bajo la GNU GENERAL PUBLIC LICENSE v3

'función arcotangente con cuadrante
Function ArcTan2(X, Y)
PI = 3.14159265358979
PI_2 = 1.5707963267949 'pi radianes
    Select Case X
        Case Is > 0
            ArcTan2 = Atn(Y / X)
        Case Is < 0
            ArcTan2 = Atn(Y / X) + PI * Sgn(Y)
            If Y = 0 Then ArcTan2 = ArcTan2 + PI
        Case Is = 0
            ArcTan2 = PI_2 * Sgn(Y)
    End Select
End Function
'función recursiva a través de las carpetas y escritura de las tablas
Function manejar_vistas(navis_doc, vistas_guardadas, i, nombre_carpeta)
Dim campoh As Double
Dim campov As Double
Dim relacion_aspecto As Double
For Each vista In vistas_guardadas
If vista.Type = eSavedViewType_Folder Then
    nombre_carpeta = vista.Name
    manejar_vistas navis_doc, vista.SavedViews, i, nombre_carpeta
Else
nombre_vista = vista.Name
escala_plano = Cells(37, 2)
navis_doc.state.CurrentView = vista.anonview
navis_doc.StayOpen
posicionx = navis_doc.state.CurrentView.ViewPoint.Camera.Position.Data1 + Cells(46, 2)
posiciony = navis_doc.state.CurrentView.ViewPoint.Camera.Position.Data2 + Cells(47, 2)
posicionz = navis_doc.state.CurrentView.ViewPoint.Camera.Position.Data3
distancia_focal = navis_doc.state.CurrentView.ViewPoint.FocalDistance
direccionx = navis_doc.state.CurrentView.ViewPoint.Camera.GetViewDir().Data1
direcciony = navis_doc.state.CurrentView.ViewPoint.Camera.GetViewDir().Data2
direccionz = navis_doc.state.CurrentView.ViewPoint.Camera.GetViewDir().Data3
posicionxprima = posicionx + distancia_focal * direccionx
posicionyprima = posiciony + distancia_focal * direcciony
posicionzprima = posicionz + distancia_focal * direccionz
cateto_adyacente = posicionxprima - posicionx
cateto_opuesto = posicionyprima - posiciony
angulo = ArcTan2(cateto_adyacente, cateto_opuesto) 'funcion arctan2(X, Y)
Cells(i, 7) = posicionx * escala_plano
Cells(i, 8) = posiciony * escala_plano
Cells(i, 9) = posicionz * escala_plano
Cells(i, 10) = posicionxprima * escala_plano
Cells(i, 11) = posicionyprima * escala_plano
Cells(i, 12) = posicionzprima * escala_plano
Cells(i, 13) = angulo * 180 / 3.14159265358979 'grados
Cells(i, 14) = angulo 'radianes
Cells(i, 15) = distancia_focal * escala_plano
Cells(i, 16) = nombre_vista
Cells(i, 19) = nombre_carpeta
relacion_aspecto = navis_doc.state.CurrentView.ViewPoint.Camera.AspectRatio
campov = navis_doc.state.CurrentView.ViewPoint.Camera.HeightField
campoh = 2 * Atn(relacion_aspecto * Tan(campov / 2)) 'calculo fov
campohgrados = campoh * 180 / 3.14159265358979
Cells(i, 17) = campohgrados
'comienza condicionales lentes
If campohgrados > 100 Then
Cells(i, 18) = "3mm"
ElseIf campohgrados > 90 Then
Cells(i, 18) = "4mm"
ElseIf campohgrados > 80 Then
Cells(i, 18) = "5mm"
ElseIf campohgrados > 72 Then
Cells(i, 18) = "6mm"
ElseIf campohgrados > 62 Then
Cells(i, 18) = "7mm"
ElseIf campohgrados > 53 Then
Cells(i, 18) = "8mm"
ElseIf campohgrados > 43 Then
Cells(i, 18) = "9mm"
ElseIf campohgrados > 33 Then
Cells(i, 18) = "10mm"
Else: Cells(i, 18) = "Error!"
End If
'acaba condicionales lentes
i = i + 1
End If
Next
End Function
'lee objeto bloque desde Autocad
Sub todos_bloques()
Dim bloques As AcadBlocks
Dim bl As AcadBlock
Set bloques = AutoCAD.Application.ActiveDocument.Blocks
i = 2
For Each bl In bloques
    If bl.IsLayout = False Then
        Cells(i, 5) = bl.Name
        i = i + 1
    End If
Next
End Sub
'Inicializar Navisworks, escoger archivo, arrancar la función de calculo
Sub obtener_navis()
Set navis_doc = CreateObject("Navisworks.Document")
navis_doc.Visible = True

'dialogo abrir archivo
Dim fDialog As FileDialog, result As Integer
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
fDialog.AllowMultiSelect = False
fDialog.Title = "Escoge un fichero"
fDialog.InitialFileName = ThisWorkbook.Path '"C:\"
fDialog.Filters.Clear
fDialog.Filters.Add "Ficheros NavisWorks", "*.nwd"
fDialog.Filters.Add "Todos los ficheros", "*.*"
If fDialog.Show = -1 Then
archivo_escogido = fDialog.SelectedItems(1)
End If
navis_doc.OpenFile (archivo_escogido)

Dim i As Integer
i = 2
Dim nombre_vista As String
Dim nombre_carpeta As String
Dim vista
Dim vistas_guardadas

manejar_vistas navis_doc, navis_doc.state.SavedViews, i, nombre_carpeta

End Sub
'Inserta bloques en AutoCAD
Sub insertar_bloques()
Dim bl As AcadBlockReference
Dim pto(0 To 2) As Double
Dim atributos As Variant
'i=2 para titulos
i = 2
escala = Cells(36, 2) 'opcion 1
Do While Cells(i, 7) <> ""
    nombre = Cells(i, 6)
    pto(0) = Cells(i, 7)
    pto(1) = Cells(i, 8)

    desplazamientoplantas = Cells(45, 2)
    For k = 0 To 6
        If Cells(i, 9) > Cells(44 - k, 2) Then 'de mayor altura a menor
        pto(1) = pto(1) + desplazamientoplantas * (6 - k) 'planta baja x0
        Exit For
        End If
    Next
    'test (punto universal erroneo)
    'pto(0) = pto(0) + 250000
    'test
    angulo = Cells(i, 14)
    Set bl = AutoCAD.Application.ActiveDocument.ModelSpace.InsertBlock(pto, nombre, escala, escala, escala, angulo)
    For Each atributos In bl.GetAttributes
    If atributos.TagString = "CAMERA_NUMBER" Then
    atributos.TextString = Left(CStr(Cells(i, 16)), 4) 'cogemos solo las 4 primeras posiciones
    ElseIf atributos.TagString = "MOUNT_HEIGHT" Then
            For j = 0 To 5
            If Cells(i, 9) > Cells(44 - j, 2) Then
            atributos.TextString = CStr((Cells(i, 9) - Cells(44 - j, 2)) / 1000) & "m" 'de mayor altura a menor
            Exit For
            Else: atributos.TextString = CStr(Cells(i, 9) / 1000) & "m"
            End If
            Next
    End If
    Next
    
i = i + 1
Loop
End Sub

'Genera informe en Word
Sub generar_informe3()
    Dim aplicacionWord As Word.Application
    Dim documentoWord As Word.Document
    Dim parrafo As Word.paragraph
    Dim imagen As Word.paragraph
    Set aplicacionWord = New Word.Application
    aplicacionWord.Visible = True
    Set documentoWord = aplicacionWord.Documents.Add()
    Dim tabla_a_copiar As Excel.Range
    Set tabla_a_copiar = ThisWorkbook.Worksheets("PROGRAMA").Range("w2:z8")
    Cells(3, "X") = Cells(38, "B") ' X3 es fija
    
    'dialogo seleccionar carpeta
Dim fDialog As FileDialog, result As Integer
Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
fDialog.AllowMultiSelect = False
fDialog.Title = "Escoge la carpeta con las imágenes extraídas de Navisworks"
fDialog.InitialFileName = ThisWorkbook.Path '"C:\"
fDialog.Filters.Clear
If fDialog.Show = -1 Then
archivo_escogido = fDialog.SelectedItems(1)
End If
    
Dim i As Integer
    For i = 2 To ActiveSheet.Cells(Rows.Count, "G").End(xlUp).Row
    nombre_pto_vista = Cells(i, 16)
    nombre_carpeta = Cells(i, 19)
    fov = Cells(i, 17)
    lente = Cells(i, 18)
    Cells(2, "X") = nombre_pto_vista ' & ": " & nombre_carpeta
    Cells(4, "X") = fov
    Cells(5, "X") = lente
            
    posx = Cells(i, 7).Value
    Cells(7, "X").Value = posx
    posy = Cells(i, 8).Value
    Cells(7, "Y").Value = posy
    posz = Cells(i, 9).Value
    Cells(7, "Z") = posz
    
    posxla = Cells(i, 10).Value
    Cells(8, "X").Value = posxla
    posyla = Cells(i, 11).Value
    Cells(8, "Y").Value = posyla
    poszla = Cells(i, 12).Value
    Cells(8, "Z").Value = poszla
    
    ThisWorkbook.Worksheets("PROGRAMA").Activate
    
    Set parrafo = documentoWord.Paragraphs.Add()
    'limpia el portapapeles
    Application.CutCopyMode = False
    tabla_a_copiar.Copy
    parrafo.Range.PasteExcelTable _
                               LinkedToExcel:=False, _
                               WordFormatting:=False, _
                               RTF:=True
                               'wordformatting 0 copia estilos de excel

    Application.CutCopyMode = False
    'limpia el portapapeles
    
    'escribir función de leading zeros
    Set imagen = documentoWord.Paragraphs.Add()
        If i - 1 < 10 Then
        imagen.Range.InlineShapes.AddPicture (archivo_escogido & "\vp000" & (i - 1) & ".jpg")
        ElseIf i - 1 < 100 Then
        imagen.Range.InlineShapes.AddPicture (archivo_escogido & "\vp00" & (i - 1) & ".jpg")
        ElseIf i - 1 < 1000 Then
        imagen.Range.InlineShapes.AddPicture (archivo_escogido & "\vp0" & (i - 1) & ".jpg")
        Else
        imagen.Range.InlineShapes.AddPicture (archivo_escogido & "\vp" & (i - 1) & ".jpg")
        End If
    imagen.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    
Set tablasWord = documentoWord.Tables(i - 1)
    tablasWord.AutoFitBehavior (wdAutoFitWindow)
    tablasWord.Range.Paragraphs.KeepWithNext = True 'mantener tablas juntas
    For j = 1 To 7 'estetica:alinea texto en tabla, tabla tiene 7 filas
    With tablasWord.Rows(j)
    .Cells.VerticalAlignment = wdAlignVerticalCenter
    .AllowBreakAcrossPages = False
    End With
    Next

Next
End Sub
Sub limpiar_campos()
ultima_fila = ActiveSheet.UsedRange.Rows.Count
Set r = Range(Cells(2, 5), Cells(ultima_fila, "S"))
r.ClearContents
End Sub
