Attribute VB_Name = "mod_BaseMarcas"
Option Explicit

'==============================================================================
' mod_BaseMarcas   -   Base_Marcas.xlsm
'
' BASE de fondos tradicionales. Una sola fuente: el libro de Marcas.
'
' Esta base SOLO CARGA Y GUARDA. No compone retornos ni arma ventanas. Eso lo
' hace el producto (PriceAction.xlsm), que abre este libro en solo lectura.
'
' Hojas:
'   Config          ruta del libro de Marcas, hoja y columnas
'   Maestro_Fondos  TU lista de fondos: codigo, nombre, categoria, fuera/REF
'   Datos_Fondos    HISTORICO que escribes tu: fecha, codigo, duracion,
'                   spread, CY. Una fila por fondo y fecha. Nunca se borra.
'   Marcas_Detalle  una marca por fondo y fecha, tal cual viene de Marcas
'   Panel           los botones
'   Log             el ultimo error, completo
'
' Rutina:  1. Cargar marcas       cuando llegan marcas nuevas
'          2. Plantilla de datos  cada mes/trimestre, cuando llega la
'                                 presentacion del gestor: te arma las filas
'                                 y tu pones los numeros
'
' Reglas:
'   - Marcas se abre SIEMPRE en solo lectura y se cierra. Nunca se escribe.
'   - Marcas_Detalle se rehace entera en cada carga (Marcas ya es el
'     historico completo). Datos_Fondos NUNCA se toca por macro.
'   - Nada se corrige en silencio: lo raro cae en la columna Flag.
'==============================================================================

Private Const VER As String = "1.0"

Private Const SH_CFG As String = "Config"
Private Const SH_MAE As String = "Maestro_Fondos"
Private Const SH_DAT As String = "Datos_Fondos"
Private Const SH_MDE As String = "Marcas_Detalle"
Private Const SH_LOG As String = "Log"
Private Const SH_PAN As String = "Panel"

' Maestro_Fondos
'   A Codigo SBS  B Nombre  C Categoria  D Fuera / REF  E Ultima marca
Private Const MA_COD As Long = 1
Private Const MA_NOM As Long = 2
Private Const MA_CAT As Long = 3
Private Const MA_FUE As Long = 4
Private Const MA_ULT As Long = 5
Private Const MA_COLS As Long = 5

' Datos_Fondos
'   A Fecha  B Codigo SBS  C Nombre  D Duracion  E Spread (pb)  F CY  G Fuente
Private Const DA_COLS As Long = 7

' Marcas_Detalle
'   A Codigo SBS  B Fecha  C Var Ajustada  D Nombre en Marcas  E Flag
Private Const MD_COLS As Long = 5

Private mT0 As Double
Private mPasos As String
Private mDetalle As String


'==============================================================================
'  CONFIG
'==============================================================================
Public Sub CrearConfig()
    Dim ws As Worksheet, prev As Object, k As Variant
    Dim scrPrev As Boolean

    scrPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Falla

    Set prev = NuevoDic()
    On Error Resume Next
    For Each k In Array("C2", "C3", "C4", "C5", "C8", "C9", "C10", "C11")
        prev(CStr(k)) = ThisWorkbook.Worksheets(SH_CFG).Range(CStr(k)).Value
    Next k
    On Error GoTo Falla

    Set ws = HojaOCrea(SH_CFG)
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    ws.Range("B1").Value = "BASE MARCAS  -  configuracion"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 13

    Cfg ws, 2, "Libro de Marcas (ruta completa)", Rec(prev, "C2", ""), _
        "Con carpeta y extension. Se abre en solo lectura."
    Cfg ws, 3, "Hoja de Marcas", Rec(prev, "C3", "Hoja 3"), "Nombre o numero de la hoja."
    Cfg ws, 4, "Fila de la cabecera", Rec(prev, "C4", 3), ""
    Cfg ws, 5, "Primera fila de datos", Rec(prev, "C5", 4), ""
    ws.Range("C2:C5").Interior.Color = RGB(255, 242, 204)

    ws.Range("B7").Value = "COLUMNAS DE MARCAS  (letra; vacio = la busca por la cabecera)"
    ws.Range("B7").Font.Bold = True
    Map ws, 8, "Columna Dia", Txt(Rec(prev, "C8", ""))
    Map ws, 9, "Columna Codigo SBS", Txt(Rec(prev, "C9", "G"))
    Map ws, 10, "Columna Var Ajustada", Txt(Rec(prev, "C10", "F"))
    Map ws, 11, "Columna Nombre del fondo", Txt(Rec(prev, "C11", ""))
    ws.Range("D8").Value = "Busca 'DIA' o 'FECHA'."
    ws.Range("D9").Value = "Busca 'CODIGO SBS'."
    ws.Range("D10").Value = "Busca 'VAR AJ' / 'AJUSTADA'."
    ws.Range("D11").Value = "Busca 'NOMBRE' o 'FONDO'. Solo informativo: el nombre que manda es el del Maestro."
    ws.Range("D8:D11").Font.Color = RGB(128, 128, 128)
    ws.Range("C8:C11").Interior.Color = RGB(255, 242, 204)
    ws.Range("C8:C11").HorizontalAlignment = xlCenter

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C").ColumnWidth = 60
    ws.Columns("D").ColumnWidth = 80

    Application.ScreenUpdating = scrPrev
    MsgBox "Config lista (v" & VER & ")." & vbCrLf & vbCrLf & _
           "1. Pon la ruta del libro de Marcas en C2." & vbCrLf & _
           "2. Crear Maestro y escribe tus fondos." & vbCrLf & _
           "3. Cargar marcas.", vbInformation, "CrearConfig"
    Exit Sub

Falla:
    Application.ScreenUpdating = True
    MsgBox "CrearConfig se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "CrearConfig"
End Sub


'==============================================================================
'  MAESTRO Y DATOS
'==============================================================================
Public Sub CrearMaestro()
    ' Crea Maestro_Fondos y Datos_Fondos, o les repone las cabeceras.
    ' NUNCA borra tus filas.
    Dim ws As Worksheet, existia As Boolean

    On Error GoTo Falla
    existia = HojaExiste(SH_MAE)
    Set ws = HojaOCrea(SH_MAE)
    CabeceraMaestro ws
    ws.Range("G1").Value = "COMO SE LLENA"
    ws.Range("G1").Font.Bold = True
    ws.Range("G2").Value = "A  Codigo SBS del fondo, el mismo que sale en Marcas."
    ws.Range("G3").Value = "B  Nombre como quieres verlo en el cuadro."
    ws.Range("G4").Value = "C  Categoria (ej. Fondos Tradicionales)."
    ws.Range("G5").Value = "D  'SI' = fuera del Total.  'REF' = se muestra pero no cuenta.  Vacio = normal."
    ws.Range("G6").Value = "E  La llena 'Cargar marcas': fecha de la ultima marca. 'NO' = no esta en Marcas."
    ws.Range("G2:G6").Font.Color = RGB(100, 100, 100)
    ws.Columns("G").ColumnWidth = 90

    Set ws = HojaOCrea(SH_DAT)
    CabeceraDatos ws
    ws.Range("I1").Value = "COMO SE LLENA"
    ws.Range("I1").Font.Bold = True
    ws.Range("I2").Value = "Una fila por fondo y fecha. Se AGREGA, nunca se borra: asi queda el historico."
    ws.Range("I3").Value = "Boton 'Plantilla de datos': te arma las filas de todos los fondos con la fecha que digas."
    ws.Range("I4").Value = "Tu pones duracion, spread (en puntos basicos: 139, no 1.39) y CY (en %: 0.0915 o 9.15%)."
    ws.Range("I5").Value = "G  De donde salio: 'Presentacion 3T26', 'Factsheet set-26'... para auditar."
    ws.Range("I6").Value = "El producto toma, para cada fondo, la fila con la ULTIMA fecha <= corte del cuadro."
    ws.Range("I2:I6").Font.Color = RGB(100, 100, 100)
    ws.Columns("I").ColumnWidth = 95

    If existia Then
        MsgBox "Maestro_Fondos y Datos_Fondos: cabeceras repuestas. Tus filas no se tocaron.", _
               vbInformation, "CrearMaestro"
    Else
        MsgBox "Hojas creadas." & vbCrLf & vbCrLf & _
               "Escribe tus fondos en Maestro_Fondos (A, B, C)." & vbCrLf & _
               "Luego: Cargar marcas  ->  Plantilla de datos.", vbInformation, "CrearMaestro"
    End If
    Exit Sub

Falla:
    MsgBox "CrearMaestro se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "CrearMaestro"
End Sub


Private Sub CabeceraMaestro(ws As Worksheet)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, MA_COLS)).Value = _
        Array("Codigo SBS", "Nombre", "Categoria", "Fuera / REF", "Ultima marca")
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, MA_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(1, MA_COD), ws.Cells(1, MA_FUE)).Interior.Color = RGB(191, 143, 0)
    ws.Columns(MA_COD).ColumnWidth = 18
    ws.Columns(MA_NOM).ColumnWidth = 36
    ws.Columns(MA_CAT).ColumnWidth = 22
    ws.Columns(MA_FUE).ColumnWidth = 11
    ws.Columns(MA_ULT).ColumnWidth = 14
    ws.Columns(MA_COD).NumberFormat = "@"
    ws.Columns(MA_ULT).NumberFormat = "dd/mm/yyyy"
    CongelarPaneles ws, 1, 2
End Sub


Private Sub CabeceraDatos(ws As Worksheet)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, DA_COLS)).Value = _
        Array("Fecha", "Codigo SBS", "Nombre", "Duracion", "Spread (pb)", "CY", "Fuente")
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, DA_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(191, 143, 0)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 36
    ws.Range(ws.Columns(4), ws.Columns(6)).ColumnWidth = 12
    ws.Columns(7).ColumnWidth = 30
    ws.Columns(1).NumberFormat = "dd/mm/yyyy"
    ws.Columns(2).NumberFormat = "@"
    ws.Columns(4).NumberFormat = "0.00"
    ws.Columns(5).NumberFormat = "0"
    ws.Columns(6).NumberFormat = "0.00%"
    CongelarPaneles ws, 1, 3
End Sub


Private Function CodigosMaestro() As Object
    ' codigo -> Array(fila, nombre)
    Dim d As Object, ws As Worksheet, i As Long, ult As Long, cod As String
    Set d = NuevoDic()
    Set CodigosMaestro = d
    If Not HojaExiste(SH_MAE) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_MAE)
    ult = ws.Cells(ws.Rows.Count, MA_COD).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, MA_COD).Value))
        If Len(cod) > 0 Then
            If Not d.Exists(cod) Then d(cod) = Array(i, Trim$(Txt(ws.Cells(i, MA_NOM).Value)))
        End If
    Next i
End Function


Public Sub PlantillaDatos()
    ' Agrega al final de Datos_Fondos una fila por fondo del Maestro con la
    ' fecha que escribas, para que solo pongas los numeros. Si ya hay filas
    ' con esa fecha para un fondo, a ese no lo repite.
    Dim ws As Worksheet, s As String, d As Double
    Dim codigos As Object, k As Variant, arr As Variant
    Dim ult As Long, i As Long, r As Long, n As Long, ya As Object

    On Error GoTo Falla
    Set codigos = CodigosMaestro()
    If codigos.Count = 0 Then
        MsgBox "Maestro_Fondos no tiene codigos en la columna A.", vbExclamation, "PlantillaDatos"
        Exit Sub
    End If
    s = InputBox("Fecha de los datos (dd/mm/aaaa)?" & vbCrLf & vbCrLf & _
                 "Normalmente el cierre del mes o trimestre de la presentacion." & vbCrLf & _
                 "Se agrega una fila por fondo; tu llenas duracion, spread y CY.", _
                 "PlantillaDatos", Format$(Date, "dd/mm/yyyy"))
    If Len(Trim$(s)) = 0 Then Exit Sub
    If Not IsDate(s) Then
        MsgBox "No entendi la fecha: " & s, vbExclamation, "PlantillaDatos"
        Exit Sub
    End If
    d = CDbl(CDate(s))

    Set ws = HojaOCrea(SH_DAT)
    CabeceraDatos ws
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 1 Then ult = 1
    Set ya = NuevoDic()
    For i = 2 To ult
        If EsNumero(ws.Cells(i, 1).Value) Then
            If CDbl(ws.Cells(i, 1).Value) = d Then ya(SinGuiones(Txt(ws.Cells(i, 2).Value))) = 1
        End If
    Next i

    r = ult
    For Each k In codigos.Keys
        If Not ya.Exists(CStr(k)) Then
            arr = codigos(k)
            r = r + 1
            ws.Cells(r, 1).Value = d
            ws.Cells(r, 2).Value = "'" & CStr(k)
            ws.Cells(r, 3).Value = CStr(arr(1))
            ws.Range(ws.Cells(r, 4), ws.Cells(r, 6)).Interior.Color = RGB(255, 242, 204)
            n = n + 1
        End If
    Next k
    ws.Activate
    If n > 0 Then ws.Cells(ult + 1, 4).Select

    MsgBox "Agregadas " & n & " filas con fecha " & Format$(d, "dd/mm/yyyy") & "." & vbCrLf & vbCrLf & _
           "Llena duracion, spread (pb) y CY en las celdas amarillas." & vbCrLf & _
           IIf(codigos.Count - n > 0, (codigos.Count - n) & " fondos ya tenian fila con esa fecha.", ""), _
           vbInformation, "PlantillaDatos"
    Exit Sub

Falla:
    MsgBox "PlantillaDatos se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "PlantillaDatos"
End Sub


'==============================================================================
'  CARGAR MARCAS  ->  Marcas_Detalle
'==============================================================================
Public Sub CargarMarcas()
    Dim wsCfg As Worksheet, wsMae As Worksheet, wsDet As Worksheet
    Dim wb As Workbook, ws As Worksheet
    Dim ruta As String, hoja As String, v As Variant
    Dim cab As Long, fil0 As Long, filaCab As Long
    Dim colDia As Long, colCod As Long, colVar As Long, colNom As Long
    Dim ultF As Long, ultC As Long, i As Long, nFilas As Long
    Dim codigos As Object, ultMar As Object, llave As Object
    Dim cod As String, dFe As Double, vAdj As Double, k As Variant, arr As Variant
    Dim det() As Variant, nDet As Long, pos As Long
    Dim nUsadas As Long, nDup As Long, raros As Long, nSin As Long
    Dim yaAbierto As Boolean, msg As String
    Dim scrPrev As Boolean, calcPrev As XlCalculation, evPrev As Boolean

    mT0 = Timer: mPasos = "": mDetalle = ""
    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Paso "Validando Config"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    ruta = Trim$(Txt(wsCfg.Range("C2").Value))
    If Len(ruta) = 0 Then Err.Raise vbObjectError + 30, , "Falta la ruta del libro de Marcas en Config C2."
    If Len(Dir(ruta)) = 0 Then Err.Raise vbObjectError + 36, , "No existe el archivo: " & ruta
    hoja = Trim$(Txt(wsCfg.Range("C3").Value))
    If Len(hoja) = 0 Then hoja = "3"
    cab = CLng(NumDef(wsCfg.Range("C4").Value, 1))
    fil0 = CLng(NumDef(wsCfg.Range("C5").Value, 2))

    Paso "Leyendo el Maestro"
    Set codigos = CodigosMaestro()
    If codigos.Count = 0 Then Err.Raise vbObjectError + 31, , _
        "Maestro_Fondos no tiene codigos en la columna A."

    Paso "Abriendo Marcas en solo lectura"
    yaAbierto = LibroYaAbierto(ruta)
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    On Error Resume Next
    If IsNumeric(hoja) Then
        Set ws = wb.Worksheets(CLng(hoja))
    Else
        Set ws = wb.Worksheets(hoja)
    End If
    On Error GoTo Falla
    If ws Is Nothing Then
        If Not yaAbierto Then wb.Close SaveChanges:=False
        Err.Raise vbObjectError + 33, , "No encontre la hoja '" & hoja & "' en Marcas."
    End If

    Paso "Leyendo las marcas a memoria"
    ultF = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ultC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If ultF < 1 Then ultF = 1
    If ultC < 1 Then ultC = 1
    v = ws.Range(ws.Cells(1, 1), ws.Cells(ultF, ultC)).Value2
    If Not yaAbierto Then wb.Close SaveChanges:=False
    Set wb = Nothing
    Set ws = Nothing
    If Not IsArray(v) Then Err.Raise vbObjectError + 34, , "La hoja de Marcas salio vacia."

    Paso "Ubicando las columnas"
    colDia = Letra(Txt(wsCfg.Range("C8").Value), 0)
    colCod = Letra(Txt(wsCfg.Range("C9").Value), 0)
    colVar = Letra(Txt(wsCfg.Range("C10").Value), 0)
    colNom = Letra(Txt(wsCfg.Range("C11").Value), 0)
    filaCab = cab
    If colDia = 0 Then colDia = BuscaCol(v, cab, "DIA|FECHA", filaCab)
    If colCod = 0 Then colCod = BuscaCol(v, cab, "CODIGO SBS|CODIGO", filaCab)
    If colVar = 0 Then colVar = BuscaCol(v, cab, "VAR AJ|VARIACION AJ|VARAJ|AJUSTADA", filaCab)
    If colNom = 0 Then colNom = BuscaCol(v, cab, "NOMBRE|FONDO", filaCab)
    If filaCab > cab Then
        cab = filaCab
        If fil0 <= cab Then fil0 = cab + 1
    End If
    If colDia = 0 Or colCod = 0 Or colVar = 0 Then
        msg = "No pude ubicar las columnas de Marcas." & vbCrLf & vbCrLf
        msg = msg & "Dia: " & LetraODash(colDia) & "   Codigo SBS: " & LetraODash(colCod) & _
              "   Var Adj: " & LetraODash(colVar) & vbCrLf & vbCrLf
        msg = msg & "Esto es lo que lei en las primeras filas de la hoja '" & hoja & "':" & vbCrLf & Cabeceras(v)
        msg = msg & vbCrLf & "Escribe las letras a mano en Config C8, C9 y C10."
        Err.Raise vbObjectError + 35, , msg
    End If

    Paso "Filtrando las marcas de tus fondos"
    ' Buffer (columna, fila): ReDim Preserve solo crece la ULTIMA dimension.
    ReDim det(1 To MD_COLS, 1 To 50000)
    nDet = 0
    Set llave = NuevoDic()    ' codigo|fecha -> posicion en det (para no duplicar)
    Set ultMar = NuevoDic()
    nFilas = UBound(v, 1)
    For i = fil0 To nFilas
        If EsNumero(v(i, colDia)) Then
            cod = SinGuiones(Txt(v(i, colCod)))
            If codigos.Exists(cod) Then
                dFe = CDbl(v(i, colDia))
                If EsNumero(v(i, colVar)) Then
                    vAdj = CDbl(v(i, colVar))
                    If llave.Exists(cod & "|" & CLng(dFe)) Then
                        ' misma marca dos veces: manda la ultima del archivo
                        pos = CLng(llave(cod & "|" & CLng(dFe)))
                        nDup = nDup + 1
                    Else
                        nDet = nDet + 1
                        If nDet > UBound(det, 2) Then ReDim Preserve det(1 To MD_COLS, 1 To UBound(det, 2) * 2)
                        pos = nDet
                        llave(cod & "|" & CLng(dFe)) = pos
                    End If
                    det(1, pos) = "'" & cod
                    det(2, pos) = dFe
                    det(3, pos) = vAdj
                    If colNom > 0 Then det(4, pos) = Trim$(Txt(v(i, colNom)))
                    det(5, pos) = ""
                    If Abs(vAdj) > 1 Then
                        det(5, pos) = "REVISAR: |Var Adj| mayor a 1"
                        raros = raros + 1
                    ElseIf vAdj <= -1 Then
                        det(5, pos) = "REVISAR: Var Adj <= -100%"
                        raros = raros + 1
                    End If
                    nUsadas = nUsadas + 1
                    If Not ultMar.Exists(cod) Then ultMar(cod) = 0
                    If dFe > CDbl(ultMar(cod)) Then ultMar(cod) = dFe
                End If
            End If
        End If
    Next i

    Paso "Escribiendo Marcas_Detalle (" & Format$(nDet, "#,##0") & " marcas)"
    Set wsDet = HojaOCrea(SH_MDE)
    EscribirDetalle wsDet, det, nDet

    Paso "Marcando la ultima marca en el Maestro"
    Set wsMae = ThisWorkbook.Worksheets(SH_MAE)
    For Each k In codigos.Keys
        arr = codigos(k)
        If ultMar.Exists(CStr(k)) Then
            wsMae.Cells(CLng(arr(0)), MA_ULT).Value = CDbl(ultMar(CStr(k)))
            wsMae.Cells(CLng(arr(0)), MA_ULT).NumberFormat = "dd/mm/yyyy"
        Else
            wsMae.Cells(CLng(arr(0)), MA_ULT).Value = "NO"
            nSin = nSin + 1
        End If
    Next k

    Restaurar calcPrev, scrPrev, evPrev

    msg = "v" & VER & " - CargarMarcas terminado en " & Format$(Timer - mT0, "0.0") & " s." & vbCrLf & vbCrLf
    msg = msg & mPasos & vbCrLf
    msg = msg & "Columnas: Dia " & NumALetra(colDia) & "  Codigo " & NumALetra(colCod) & _
          "  Var Adj " & NumALetra(colVar) & "  Nombre " & LetraODash(colNom) & vbCrLf
    msg = msg & "Filas en Marcas: " & Format$(nFilas - fil0 + 1, "#,##0") & vbCrLf
    msg = msg & "Marcas de tus fondos: " & Format$(nUsadas, "#,##0") & vbCrLf
    If nDup > 0 Then msg = msg & "Repetidas (misma fecha y fondo, se quedo la ultima): " & nDup & vbCrLf
    If raros > 0 Then msg = msg & "REVISAR: " & raros & " marcas raras, mira la columna Flag." & vbCrLf
    msg = msg & "Fondos del Maestro: " & codigos.Count
    If nSin > 0 Then msg = msg & "   (" & nSin & " sin ninguna marca: columna E = NO)"
    msg = msg & vbCrLf & vbCrLf & "Marcas_Detalle queda con " & Format$(nDet, "#,##0") & " filas."
    Aviso msg, "CargarMarcas"
    Exit Sub

Falla:
    Reventar wb, calcPrev, scrPrev, evPrev, "CargarMarcas"
End Sub


Private Sub EscribirDetalle(ws As Worksheet, det As Variant, nDet As Long)
    Dim out() As Variant, i As Long, j As Long
    ws.Cells.Clear
    ws.Range(ws.Cells(1, 1), ws.Cells(1, MD_COLS)).Value = _
        Array("Codigo SBS", "Fecha", "Var Ajustada", "Nombre en Marcas", "Flag")
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, MD_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
    End With
    ws.Columns(1).ColumnWidth = 18
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 14
    ws.Columns(4).ColumnWidth = 36
    ws.Columns(5).ColumnWidth = 34
    If nDet = 0 Then Exit Sub

    ReDim out(1 To nDet, 1 To MD_COLS)
    For i = 1 To nDet
        For j = 1 To MD_COLS
            out(i, j) = det(j, i)
        Next j
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(nDet + 1, MD_COLS)).Value = out
    ws.Range(ws.Cells(2, 2), ws.Cells(nDet + 1, 2)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 3), ws.Cells(nDet + 1, 3)).NumberFormat = "0.0000%"
    ' ordenada por codigo y fecha, para que se pueda leer de corrido
    ws.Range(ws.Cells(1, 1), ws.Cells(nDet + 1, MD_COLS)).Sort _
        Key1:=ws.Cells(2, 1), Order1:=xlAscending, _
        Key2:=ws.Cells(2, 2), Order2:=xlAscending, Header:=xlYes
    CongelarPaneles ws, 1, 2
End Sub


'==============================================================================
'  DIAGNOSTICO
'==============================================================================
Public Sub Diagnostico()
    Dim wsCfg As Worksheet, ws As Worksheet, msg As String
    Dim codigos As Object, k As Variant, arr As Variant
    Dim ult As Long, i As Long, nNo As Long, nDat As Object, dMax As Double
    Dim ruta As String

    On Error GoTo Falla
    msg = "BASE MARCAS  v" & VER & vbCrLf & String$(44, "-") & vbCrLf
    If Not HojaExiste(SH_CFG) Then
        MsgBox msg & "No existe Config. Corre Crear Config.", vbExclamation, "Diagnostico"
        Exit Sub
    End If
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    ruta = Trim$(Txt(wsCfg.Range("C2").Value))
    msg = msg & "Libro de Marcas: " & IIf(Len(ruta) > 0, ruta, "VACIO (Config C2)") & vbCrLf
    If Len(ruta) > 0 Then msg = msg & "   " & IIf(Len(Dir(ruta)) > 0, "existe", "NO EXISTE en esa ruta") & vbCrLf
    msg = msg & vbCrLf

    Set codigos = CodigosMaestro()
    msg = msg & "MAESTRO_FONDOS: " & codigos.Count & " fondos" & vbCrLf
    If HojaExiste(SH_MAE) Then
        Set ws = ThisWorkbook.Worksheets(SH_MAE)
        For Each k In codigos.Keys
            arr = codigos(k)
            If UCase$(Txt(ws.Cells(CLng(arr(0)), MA_ULT).Value)) = "NO" Then nNo = nNo + 1
        Next k
        If nNo > 0 Then msg = msg & "   " & nNo & " sin ninguna marca (columna E = NO)" & vbCrLf
    End If

    If HojaExiste(SH_MDE) Then
        Set ws = ThisWorkbook.Worksheets(SH_MDE)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        msg = msg & "MARCAS_DETALLE: " & Format$(IIf(ult > 1, ult - 1, 0), "#,##0") & " marcas" & vbCrLf
        For i = 2 To ult
            If EsNumero(ws.Cells(i, 2).Value) Then
                If CDbl(ws.Cells(i, 2).Value) > dMax Then dMax = CDbl(ws.Cells(i, 2).Value)
            End If
        Next i
        If dMax > 0 Then msg = msg & "   ultima marca: " & Format$(CDate(dMax), "dd/mm/yyyy") & vbCrLf
    Else
        msg = msg & "MARCAS_DETALLE: no existe. Corre Cargar marcas." & vbCrLf
    End If

    If HojaExiste(SH_DAT) Then
        Set ws = ThisWorkbook.Worksheets(SH_DAT)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Set nDat = NuevoDic()
        dMax = 0
        For i = 2 To ult
            If EsNumero(ws.Cells(i, 1).Value) Then
                nDat(CLng(ws.Cells(i, 1).Value)) = 1
                If CDbl(ws.Cells(i, 1).Value) > dMax Then dMax = CDbl(ws.Cells(i, 1).Value)
            End If
        Next i
        msg = msg & "DATOS_FONDOS: " & Format$(IIf(ult > 1, ult - 1, 0), "#,##0") & " filas, " & _
              nDat.Count & " fechas"
        If dMax > 0 Then msg = msg & ", ultima " & Format$(CDate(dMax), "dd/mm/yyyy")
        msg = msg & vbCrLf
        nNo = 0
        For i = 2 To ult
            If EsNumero(ws.Cells(i, 1).Value) Then
                If Not EsNumero(ws.Cells(i, 4).Value) Or Not EsNumero(ws.Cells(i, 5).Value) _
                   Or Not EsNumero(ws.Cells(i, 6).Value) Then nNo = nNo + 1
            End If
        Next i
        If nNo > 0 Then msg = msg & "   " & nNo & " filas con duracion, spread o CY vacios" & vbCrLf
    Else
        msg = msg & "DATOS_FONDOS: no existe. Corre Crear Maestro." & vbCrLf
    End If

    MsgBox msg, vbInformation, "Diagnostico"
    Exit Sub

Falla:
    MsgBox "Diagnostico se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Diagnostico"
End Sub


'==============================================================================
'  PANEL
'==============================================================================
Public Sub CrearPanel()
    Dim ws As Worksheet, i As Long, rutina As Variant, extra As Variant
    Dim scrPrev As Boolean

    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set ws = HojaOCrea(SH_PAN)
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    ws.Range("B2").Value = "BASE MARCAS  -  fondos tradicionales"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 16
    ws.Range("B2").Font.Color = RGB(192, 0, 0)
    ws.Range("B3").Value = "v" & VER & "  -  esta base solo carga y guarda. El cuadro se arma en PriceAction."
    ws.Range("B3").Font.Color = RGB(128, 128, 128)

    ws.Range("B5").Value = "RUTINA"
    ws.Range("B5").Font.Bold = True
    ws.Range("F5").Value = "CUANDO HACE FALTA"
    ws.Range("F5").Font.Bold = True

    rutina = Array( _
        "CargarMarcas", "1. Cargar marcas", _
        "PlantillaDatos", "2. Plantilla de datos")

    extra = Array( _
        "NormalizarExcel", "ARREGLAR EXCEL", _
        "CrearConfig", "Crear Config", _
        "CrearMaestro", "Crear Maestro", _
        "Diagnostico", "Diagnostico", _
        "Version", "Version")

    BotonesEnColumna ws, rutina, 60, 110, 170, 30, True
    BotonesEnColumna ws, extra, 300, 110, 170, 26, False

    ws.Range("B22").Value = "1 cuando llegan marcas nuevas. Rehace Marcas_Detalle completa."
    ws.Range("B23").Value = "2 cada mes o trimestre: arma las filas en Datos_Fondos y tu pones duracion, spread y CY."
    ws.Range("B24").Value = "Datos_Fondos nunca se borra por macro: es tu historico."
    ws.Range("B25").Value = "Si cortaste una macro con Esc: ARREGLAR EXCEL antes de seguir."
    ws.Range("B22:B25").Font.Color = RGB(128, 128, 128)

    ws.Columns("A").ColumnWidth = 2
    ws.Rows(2).RowHeight = 22
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = scrPrev
    MsgBox "Panel listo.", vbInformation, "CrearPanel"
    Exit Sub

Falla:
    Application.ScreenUpdating = True
    MsgBox "CrearPanel se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "CrearPanel"
End Sub


Private Sub BotonesEnColumna(ws As Worksheet, lista As Variant, izq As Double, _
                             arriba As Double, ancho As Double, alto As Double, _
                             destacar As Boolean)
    Dim i As Long, sh As Shape, y As Double
    y = arriba
    For i = LBound(lista) To UBound(lista) Step 2
        Set sh = ws.Shapes.AddFormControl(xlButtonControl, izq, y, ancho, alto)
        sh.OnAction = CStr(lista(i))
        sh.TextFrame.Characters.Text = CStr(lista(i + 1))
        sh.Name = "btn_" & CStr(lista(i))
        If destacar Then
            sh.TextFrame.Characters.Font.Bold = True
            sh.TextFrame.Characters.Font.Size = 10
        Else
            sh.TextFrame.Characters.Font.Size = 9
        End If
        y = y + alto + 6
    Next i
End Sub


Public Sub NormalizarExcel()
    Dim antes As String
    antes = "Calculo: " & IIf(Application.Calculation = xlCalculationAutomatic, "automatico", "MANUAL")
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox "Excel normalizado." & vbCrLf & vbCrLf & "Como estaba:  " & antes & vbCrLf & vbCrLf & _
           "Ahora: calculo automatico, pantalla y eventos activos.", vbInformation, "NormalizarExcel"
End Sub


Public Sub Version()
    MsgBox "mod_BaseMarcas  v" & VER, vbInformation, "Version"
End Sub


'==============================================================================
'  LECTURA DE MARCAS
'==============================================================================
Private Function BuscaCol(v As Variant, cab As Long, patrones As String, _
                          ByRef filaCab As Long) As Long
    ' Busca la columna cuya cabecera contenga alguno de los patrones. Empieza
    ' por la fila que dice Config y, si no la encuentra, barre las 15 primeras.
    Dim i As Long, j As Long, t As String, p As Variant, hasta As Long
    If Not IsArray(v) Then Exit Function
    hasta = UBound(v, 1)
    If hasta > 15 Then hasta = 15
    For Each p In Split(patrones, "|")
        For i = 1 To hasta
            For j = LBound(v, 2) To UBound(v, 2)
                t = NormTxt(Txt(v(i, j)))
                If Len(t) > 0 Then
                    If InStr(1, t, CStr(p), vbBinaryCompare) > 0 Then
                        BuscaCol = j
                        If i > filaCab Then filaCab = i
                        Exit Function
                    End If
                End If
            Next j
        Next i
    Next p
End Function


Private Function Cabeceras(v As Variant) As String
    Dim i As Long, j As Long, t As String, linea As String, out As String
    Dim nFil As Long, hasta As Long
    If Not IsArray(v) Then Exit Function
    hasta = UBound(v, 1)
    If hasta > 12 Then hasta = 12
    For i = 1 To hasta
        linea = ""
        For j = LBound(v, 2) To UBound(v, 2)
            t = Trim$(Txt(v(i, j)))
            If Len(t) > 0 Then
                If Len(linea) < 90 Then linea = linea & NumALetra(j) & "=" & Left$(t, 14) & "  "
            End If
        Next j
        If Len(linea) > 0 Then
            nFil = nFil + 1
            out = out & "fila " & i & ":  " & linea & vbCrLf
            If nFil >= 5 Then Exit For
        End If
    Next i
    Cabeceras = out
End Function


Private Function LibroYaAbierto(ruta As String) As Boolean
    Dim wb As Workbook, nom As String
    nom = Mid$(ruta, InStrRev(ruta, Application.PathSeparator) + 1)
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nom, vbTextCompare) = 0 Then
            LibroYaAbierto = True
            Exit Function
        End If
    Next wb
End Function


Private Function LetraODash(c As Long) As String
    If c > 0 Then
        LetraODash = NumALetra(c)
    Else
        LetraODash = "NO"
    End If
End Function


'==============================================================================
'  UTILIDADES
'==============================================================================
Private Sub Aviso(texto As String, titulo As String)
    MsgBox texto, vbInformation, titulo
End Sub


Private Sub Paso(texto As String)
    mPasos = mPasos & "  - " & texto & " (" & Format$(Timer - mT0, "0.0") & " s)" & vbCrLf
    mDetalle = texto
    Application.StatusBar = "Base Marcas: " & texto & "..."
    DoEvents
End Sub


Private Sub Restaurar(calcPrev As XlCalculation, scrPrev As Boolean, evPrev As Boolean)
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = scrPrev
    Application.EnableEvents = evPrev
    Application.StatusBar = False
End Sub


Private Sub Reventar(wb As Workbook, calcPrev As XlCalculation, scrPrev As Boolean, _
                     evPrev As Boolean, quien As String)
    Dim msg As String, nErr As Long, codigo As String, texto As String
    nErr = Err.Number
    msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False

    codigo = CodigoError(nErr) & " @ " & mDetalle
    texto = quien & " se detuvo." & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & _
            "Paso a paso:" & vbCrLf & mPasos
    GuardarLog "CODIGO: " & codigo & vbCrLf & vbCrLf & texto
    MsgBox "ESCRIBEME ESTA LINEA:" & vbCrLf & vbCrLf & "   " & codigo & vbCrLf & vbCrLf & _
           String$(40, "-") & vbCrLf & msg & vbCrLf & _
           "(el detalle completo quedo en la hoja Log)", vbCritical, quien
End Sub


Private Function CodigoError(n As Long) As String
    If n < 0 Then
        CodigoError = "E" & Format$(n - vbObjectError, "00")
    Else
        CodigoError = CStr(n)
    End If
End Function


Private Sub GuardarLog(texto As String)
    Dim ws As Worksheet, partes As Variant, i As Long
    On Error Resume Next
    Set ws = HojaOCrea(SH_LOG)
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Range("A1").Value = "Ultimo error - " & Format$(Now, "dd/mm/yyyy hh:mm:ss")
        ws.Range("A1").Font.Bold = True
        partes = Split(texto, vbCrLf)
        For i = LBound(partes) To UBound(partes)
            ws.Cells(i + 3, 1).Value = "'" & partes(i)
        Next i
        ws.Columns("A").ColumnWidth = 110
    End If
    On Error GoTo 0
End Sub


Private Sub Cfg(ws As Worksheet, fila As Long, etiqueta As String, valor As Variant, nota As String)
    ws.Cells(fila, 2).Value = etiqueta
    If Not IsEmpty(valor) Then ws.Cells(fila, 3).Value = valor
    ws.Cells(fila, 4).Value = nota
    ws.Cells(fila, 4).Font.Color = RGB(128, 128, 128)
End Sub


Private Sub Map(ws As Worksheet, fila As Long, campo As String, valor As String)
    ws.Cells(fila, 2).Value = campo
    ws.Cells(fila, 3).Value = valor
End Sub


Private Function Rec(prev As Object, celda As String, porDefecto As Variant) As Variant
    If prev.Exists(celda) Then
        If Not IsEmpty(prev(celda)) Then
            If Len(Trim$(Txt(prev(celda)))) > 0 Then
                Rec = prev(celda)
                Exit Function
            End If
        End If
    End If
    Rec = porDefecto
End Function


Private Function NuevoDic() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set NuevoDic = d
End Function


Private Function HojaOCrea(nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
    End If
    Set HojaOCrea = ws
End Function


Private Function HojaExiste(nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    HojaExiste = Not ws Is Nothing
End Function


Private Sub CongelarPaneles(ws As Worksheet, filas As Long, cols As Long)
    Dim wsPrev As Object
    On Error Resume Next
    ThisWorkbook.Activate
    Set wsPrev = ActiveSheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = filas
    ActiveWindow.SplitColumn = cols
    ActiveWindow.FreezePanes = True
    If Not wsPrev Is Nothing Then wsPrev.Activate
    On Error GoTo 0
End Sub


Private Function Letra(ByVal s As String, porDefecto As Long) As Long
    Dim i As Long, n As Long, ch As Long
    s = UCase$(Trim$(s))
    If Len(s) = 0 Then
        Letra = porDefecto
        Exit Function
    End If
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1)) - 64
        If ch < 1 Or ch > 26 Then
            Letra = porDefecto
            Exit Function
        End If
        n = n * 26 + ch
    Next i
    Letra = n
End Function


Private Function NumALetra(ByVal n As Long) As String
    Dim s As String, r As Long
    Do While n > 0
        r = ((n - 1) Mod 26)
        s = Chr$(65 + r) & s
        n = (n - r - 1) \ 26
    Loop
    NumALetra = s
End Function


Private Function SinGuiones(s As String) As String
    SinGuiones = Replace(Replace(Trim$(s), "-", ""), " ", "")
End Function


Private Function NormTxt(ByVal t As String) As String
    Dim out As String
    out = UCase$(Trim$(t))
    out = Replace(out, Chr$(193), "A")
    out = Replace(out, Chr$(201), "E")
    out = Replace(out, Chr$(205), "I")
    out = Replace(out, Chr$(211), "O")
    out = Replace(out, Chr$(218), "U")
    out = Replace(out, Chr$(220), "U")
    out = Replace(out, Chr$(209), "N")
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    NormTxt = out
End Function


Private Function Txt(v As Variant) As String
    If IsError(v) Then
        Txt = ""
    ElseIf IsEmpty(v) Then
        Txt = ""
    ElseIf IsNull(v) Then
        Txt = ""
    Else
        Txt = CStr(v)
    End If
End Function


Private Function EsNumero(v As Variant) As Boolean
    If IsError(v) Then
        EsNumero = False
    ElseIf IsEmpty(v) Then
        EsNumero = False
    ElseIf VarType(v) = vbDate Then
        EsNumero = True
    ElseIf VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then
            EsNumero = False
        Else
            EsNumero = IsNumeric(v)
        End If
    Else
        EsNumero = IsNumeric(v)
    End If
End Function


Private Function NumDef(v As Variant, porDefecto As Double) As Double
    If EsNumero(v) Then
        NumDef = CDbl(v)
    Else
        NumDef = porDefecto
    End If
End Function
