Attribute VB_Name = "mod_BaseRF"
Option Explicit

'==============================================================================
' mod_BaseRF   -   Base_RF.xlsm
'
' BASE de renta fija iliquida: cuasi soberanos, iliquidos y coinversiones.
' Una sola fuente: el vector de precios SBS ("AAAAMMDD RFL.xls").
'
' Esta base SOLO CARGA Y GUARDA. No calcula retornos, ni ventanas, ni
' nombres finales. Eso lo hace el producto (PriceAction.xlsm), que abre este
' libro en solo lectura y se lleva lo que necesita.
'
' Hojas:
'   Config        carpeta del vector, inicio del historico, mapa de columnas
'   Maestro       TU lista de codigos. Tu escribes A:D (codigo, nombre,
'                 categoria, fuera/REF). El vector llena E:N.
'   BD_Precios    un precio por codigo y dia. Crece sola, nunca se borra.
'   Columnas_Vector  ayuda: que hay en cada columna del vector
'   Panel         los botones
'   Log           el ultimo error, completo
'
' Rutina del dia:   1. Cargar vectores     (todos los dias)
'                   2. Datos del vector    (cuando cambie algo del maestro)
'
' Reglas:
'   - El vector se abre SIEMPRE en solo lectura y se cierra. Nunca se escribe.
'   - El vector pisa solo lo que trae. Lo que no trae, no se toca.
'   - Un codigo nuevo en el Maestro se carga desde el inicio del historico.
'   - Nada se corrige en silencio: lo raro cae en la columna Flag.
'==============================================================================

Private Const VER As String = "1.0"

Private Const SH_CFG As String = "Config"
Private Const SH_MAE As String = "Maestro"
Private Const SH_PRE As String = "BD_Precios"
Private Const SH_CVE As String = "Columnas_Vector"
Private Const SH_LOG As String = "Log"
Private Const SH_PAN As String = "Panel"

' Maestro: columnas
'   A Codigo SBS (tu)   B Nombre (tu)   C Categoria (tu)   D Fuera / REF (tu)
'   E Nemonico  F Emisor  G Tasa cupon  H Vencimiento  I Moneda
'   J ISIN  K Rating  L Duracion  M Spread (pb)  N Visto en el vector
Private Const MA_COD As Long = 1
Private Const MA_NOM As Long = 2
Private Const MA_CAT As Long = 3
Private Const MA_FUE As Long = 4
Private Const MA_NEM As Long = 5
Private Const MA_EMI As Long = 6
Private Const MA_CUP As Long = 7
Private Const MA_VEN As Long = 8
Private Const MA_MON As Long = 9
Private Const MA_ISI As Long = 10
Private Const MA_RAT As Long = 11
Private Const MA_DUR As Long = 12
Private Const MA_SPR As Long = 13
Private Const MA_VIS As Long = 14
Private Const MA_COLS As Long = 14

' BD_Precios: columnas
'   1 Fecha  2 Codigo SBS  3 Nemonico  4 Moneda  5 Precio limpio %
'   6 Interes acum  7 Tasa cupon  8 Vencimiento  9 TIR %  10 Flag
Private Const PR_COLS As Long = 10

Private mT0 As Double
Private mPasos As String
Private mDetalle As String
Private mIdxCarpeta As String
Private mIdxVec As Object
Private mIdxStamp As Double


'==============================================================================
'  CONFIG
'==============================================================================
Public Sub CrearConfig()
    ' Rehace la hoja Config conservando lo que ya tengas escrito.
    Dim ws As Worksheet, prev As Object, k As Variant
    Dim scrPrev As Boolean

    scrPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Falla

    Set prev = NuevoDic()
    On Error Resume Next
    For Each k In Array("C2", "C3", "C4", "C5", _
                        "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", _
                        "C16", "C17", "C18", "C19", "C20", "C21")
        prev(CStr(k)) = ThisWorkbook.Worksheets(SH_CFG).Range(CStr(k)).Value
    Next k
    On Error GoTo Falla

    Set ws = HojaOCrea(SH_CFG)
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    ws.Range("B1").Value = "BASE RF  -  configuracion"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 13

    Cfg ws, 2, "Carpeta del vector SBS", Rec(prev, "C2", ""), _
        "Donde estan los archivos 'AAAAMMDD RFL.xls'."
    Cfg ws, 3, "Inicio del historico", Rec(prev, "C3", DateSerial(2025, 10, 31)), _
        "Primera fecha que se carga. Un codigo nuevo en el Maestro se trae desde aqui."
    Cfg ws, 4, "Cargar HASTA", Rec(prev, "C4", Empty), _
        "Vacio = hasta el ultimo vector que haya en la carpeta."
    Cfg ws, 5, "Dias sin vector antes de REVISAR", Rec(prev, "C5", 5), _
        "Fines de semana y feriados se saltan solos. Mas dias que esto marca la fila."
    ws.Range("C2:C5").Interior.Color = RGB(255, 242, 204)
    ws.Range("C3:C4").NumberFormat = "dd/mm/yyyy"

    ws.Range("B7").Value = "MAPA DEL VECTOR  (letra de la columna)"
    ws.Range("B7").Font.Bold = True
    ws.Range("D7").Value = "Si no sabes una letra: boton 'Ver columnas del vector'."
    ws.Range("D7").Font.Color = RGB(128, 128, 128)
    Map ws, 8, "Codigo SBS (con guiones)", Txt(Rec(prev, "C8", "A"))
    Map ws, 9, "Nemonico", Txt(Rec(prev, "C9", "C"))
    Map ws, 10, "Emisor", Txt(Rec(prev, "C10", "E"))
    Map ws, 11, "Moneda", Txt(Rec(prev, "C11", "F"))
    Map ws, 12, "Precio limpio %", Txt(Rec(prev, "C12", "J"))
    Map ws, 13, "Interes acumulado", Txt(Rec(prev, "C13", "M"))
    Map ws, 14, "TIR %", Txt(Rec(prev, "C14", "N"))
    Map ws, 15, "Fecha de vencimiento", Txt(Rec(prev, "C15", "Q"))
    Map ws, 16, "Tasa cupon", Txt(Rec(prev, "C16", "R"))
    Map ws, 17, "Duracion", Txt(Rec(prev, "C17", "X"))
    Map ws, 18, "Rating", Txt(Rec(prev, "C18", "U"))
    Map ws, 19, "ISIN", Txt(Rec(prev, "C19", ""))
    Map ws, 20, "Spread", Txt(Rec(prev, "C20", ""))
    Cfg ws, 21, "Spread: factor a puntos basicos", Rec(prev, "C21", 100), _
        "El vector trae 1.39 -> 100.  Trae 0.0139 -> 10000.  Ya trae 139 -> 1."
    ws.Range("D19").Value = "Vacio = la columna ISIN del Maestro no se toca."
    ws.Range("D20").Value = "Vacio = la columna Spread del Maestro no se toca."
    ws.Range("D19:D20").Font.Color = RGB(128, 128, 128)
    ws.Range("C8:C21").Interior.Color = RGB(255, 242, 204)
    ws.Range("C8:C20").HorizontalAlignment = xlCenter

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 80

    Application.ScreenUpdating = scrPrev
    MsgBox "Config lista (v" & VER & ")." & vbCrLf & vbCrLf & _
           "1. Pon la carpeta del vector en C2." & vbCrLf & _
           "2. Si no sabes la letra de ISIN o Spread: 'Ver columnas del vector'." & vbCrLf & _
           "3. Crear Maestro y escribe tus codigos.", vbInformation, "CrearConfig"
    Exit Sub

Falla:
    Application.ScreenUpdating = True
    MsgBox "CrearConfig se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "CrearConfig"
End Sub


'==============================================================================
'  MAESTRO
'==============================================================================
Public Sub CrearMaestro()
    ' Crea la hoja Maestro o le repone las cabeceras. NUNCA borra tus filas.
    Dim ws As Worksheet, existia As Boolean

    On Error GoTo Falla
    existia = HojaExiste(SH_MAE)
    Set ws = HojaOCrea(SH_MAE)
    CabeceraMaestro ws

    ws.Range("P1").Value = "COMO SE LLENA"
    ws.Range("P1").Font.Bold = True
    ws.Range("P2").Value = "A  Codigo SBS, con o sin guiones. Es la llave de todo."
    ws.Range("P3").Value = "B  Nombre corto del riesgo real (ej. 'Grupo Patio'). Sin cupon ni fecha: eso lo pega el producto."
    ws.Range("P4").Value = "C  Categoria: Cuasi Soberanos / Iliquidos / Coinversiones."
    ws.Range("P5").Value = "D  'SI' = fuera del Total.  'REF' = se muestra pero no cuenta.  Vacio = normal."
    ws.Range("P6").Value = "E:N  Las llena 'Datos del vector'. No las escribas a mano: se pisan."
    ws.Range("P7").Value = "N  Fecha del ultimo vector donde aparecio el codigo. 'NO' = no esta en el vector."
    ws.Range("P2:P7").Font.Color = RGB(100, 100, 100)
    ws.Columns("P").ColumnWidth = 95

    If existia Then
        MsgBox "Maestro: cabeceras repuestas. Tus filas no se tocaron.", vbInformation, "CrearMaestro"
    Else
        MsgBox "Maestro creado." & vbCrLf & vbCrLf & _
               "Escribe tus codigos en A, el nombre en B y la categoria en C." & vbCrLf & _
               "Luego: Cargar vectores  ->  Datos del vector.", vbInformation, "CrearMaestro"
    End If
    Exit Sub

Falla:
    MsgBox "CrearMaestro se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "CrearMaestro"
End Sub


Private Sub CabeceraMaestro(ws As Worksheet)
    Dim cab As Variant, i As Long
    cab = Array("Codigo SBS", "Nombre", "Categoria", "Fuera / REF", _
                "Nemonico", "Emisor", "Tasa cupon", "Vencimiento", "Moneda", _
                "ISIN", "Rating", "Duracion", "Spread (pb)", "Visto en el vector")
    For i = 0 To MA_COLS - 1
        ws.Cells(1, i + 1).Value = cab(i)
    Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, MA_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
        .HorizontalAlignment = xlCenter
    End With
    ' lo tuyo en amarillo claro, lo del vector en gris claro
    ws.Range(ws.Cells(1, MA_COD), ws.Cells(1, MA_FUE)).Interior.Color = RGB(191, 143, 0)
    ws.Columns(MA_COD).ColumnWidth = 18
    ws.Columns(MA_NOM).ColumnWidth = 30
    ws.Columns(MA_CAT).ColumnWidth = 20
    ws.Columns(MA_FUE).ColumnWidth = 11
    ws.Columns(MA_NEM).ColumnWidth = 16
    ws.Columns(MA_EMI).ColumnWidth = 30
    ws.Range(ws.Columns(MA_CUP), ws.Columns(MA_SPR)).ColumnWidth = 12
    ws.Columns(MA_VIS).ColumnWidth = 18
    ws.Columns(MA_COD).NumberFormat = "@"
    ws.Columns(MA_VEN).NumberFormat = "dd/mm/yyyy"
    ws.Columns(MA_VIS).NumberFormat = "dd/mm/yyyy"
    ws.Columns(MA_CUP).NumberFormat = "0.000"
    ws.Columns(MA_DUR).NumberFormat = "0.00"
    ws.Columns(MA_SPR).NumberFormat = "0"
    CongelarPaneles ws, 1, 2
End Sub


Private Function CodigosMaestro() As Object
    ' codigo -> fila del Maestro. Sin guiones, igual que en BD_Precios.
    Dim d As Object, ws As Worksheet, i As Long, ult As Long, cod As String
    Set d = NuevoDic()
    Set CodigosMaestro = d
    If Not HojaExiste(SH_MAE) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_MAE)
    ult = ws.Cells(ws.Rows.Count, MA_COD).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, MA_COD).Value))
        If Len(cod) > 0 Then
            If Not d.Exists(cod) Then d(cod) = i
        End If
    Next i
End Function


'==============================================================================
'  CARGAR VECTORES  ->  BD_Precios
'==============================================================================
Public Sub CargarVectores()
    Dim wsCfg As Worksheet, wsPre As Worksheet
    Dim carpeta As String, ruta As String
    Dim dDesde As Double, dHasta As Double, dFe As Double, dIni As Double
    Dim codigos As Object, vMap As Object, nuevos As Object, ya As Object
    Dim buf() As Variant, nBuf As Long, capa As Long
    Dim nArch As Long, nSalt As Long, nFilas As Long, nHueco As Long, nYa As Long, nNue As Long
    Dim ultFecha As Double, tolDias As Long
    Dim scrPrev As Boolean, calcPrev As XlCalculation, evPrev As Boolean, alertPrev As Boolean
    Dim t As Double, msg As String

    mT0 = Timer: mPasos = "": mDetalle = ""
    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    alertPrev = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Paso "Validando Config"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    carpeta = SinBarra(Txt(wsCfg.Range("C2").Value))
    If Len(carpeta) = 0 Then Err.Raise vbObjectError + 10, , "Falta la carpeta del vector en Config C2."
    If Not IsDate(wsCfg.Range("C3").Value) Then Err.Raise vbObjectError + 12, , _
        "Falta el inicio del historico en Config C3."
    tolDias = CLng(NumDef(wsCfg.Range("C5").Value, 5))
    Set vMap = MapaVector(wsCfg)

    Paso "Leyendo los codigos del Maestro"
    Set codigos = CodigosMaestro()
    If codigos.Count = 0 Then Err.Raise vbObjectError + 11, , _
        "El Maestro no tiene codigos en la columna A." & vbCrLf & _
        "Escribe ahi los codigos SBS que quieres seguir."

    Set wsPre = HojaOCrea(SH_PRE)
    PrepararPrecios wsPre
    Set ya = FechasCargadas(wsPre, ultFecha)

    Paso "Buscando codigos sin historico"
    Set nuevos = CodigosNuevos(codigos)

    ' --- rango a cargar ---
    dIni = CDbl(CDate(wsCfg.Range("C3").Value))
    If ultFecha > 0 Then
        dDesde = ultFecha + 1
    Else
        dDesde = dIni
    End If
    ' Un codigo recien agregado al Maestro no tiene ni una fila: se vuelve al
    ' inicio del historico solo por el. Las fechas que ya estan se releen
    ' trayendo unicamente a los nuevos, para no duplicar.
    If nuevos.Count > 0 Then
        If dIni < dDesde Then dDesde = dIni
    End If

    If IsDate(wsCfg.Range("C4").Value) Then
        dHasta = CDbl(CDate(wsCfg.Range("C4").Value))
    Else
        dHasta = UltimoVector(carpeta)
    End If

    If dHasta < dDesde Then
        Restaurar calcPrev, scrPrev, evPrev
        Application.DisplayAlerts = alertPrev
        Aviso "No hay nada que cargar: BD_Precios ya llega al " & _
              Format$(CDate(ultFecha), "dd/mm/yyyy") & _
              " y ese es el ultimo vector de la carpeta." & vbCrLf & vbCrLf & _
              "Si copiaste vectores nuevos con el libro abierto: 'Refrescar indice'.", "CargarVectores"
        Exit Sub
    End If

    Paso "Leyendo vectores del " & Format$(CDate(dDesde), "dd/mm/yyyy") & _
         " al " & Format$(CDate(dHasta), "dd/mm/yyyy")

    capa = 60000
    ReDim buf(1 To capa, 1 To PR_COLS)
    nBuf = 0
    nHueco = 0

    For dFe = dDesde To dHasta
        If ya.Exists(CLng(dFe)) Then
            ' Fecha ya cargada = ese dia hubo vector: corta la racha de huecos.
            nHueco = 0
            If nuevos.Count = 0 Then
                nYa = nYa + 1
            Else
                ruta = RutaVector(carpeta, dFe)
                If Len(ruta) = 0 Then
                    nSalt = nSalt + 1
                Else
                    mDetalle = "Relectura " & Format$(CDate(dFe), "dd/mm/yyyy")
                    Application.StatusBar = "Codigos nuevos, " & _
                        Format$(CDate(dFe), "dd/mm/yyyy") & "   (" & nNue & " filas, " & _
                        Format$(Timer - mT0, "0") & " s)"
                    nNue = nNue + LeerVector(ruta, dFe, vMap, nuevos, buf, nBuf, capa, "")
                    nArch = nArch + 1
                    If nBuf > 40000 Or (nArch Mod 20) = 0 Then
                        VolcarPrecios wsPre, buf, nBuf
                        nBuf = 0
                    End If
                    DoEvents
                End If
            End If
        Else
            ruta = RutaVector(carpeta, dFe)
            If Len(ruta) = 0 Then
                nSalt = nSalt + 1
                nHueco = nHueco + 1
            Else
                mDetalle = "Vector " & Format$(CDate(dFe), "dd/mm/yyyy")
                Application.StatusBar = "Vector " & Format$(CDate(dFe), "dd/mm/yyyy") & _
                    "   (" & nArch + 1 & " archivos, " & nFilas & " filas, " & _
                    Format$(Timer - mT0, "0") & " s)"
                nFilas = nFilas + LeerVector(ruta, dFe, vMap, codigos, buf, nBuf, capa, _
                                             IIf(nHueco > tolDias, "REVISAR: " & nHueco & _
                                                 " dias sin vector antes de esta fecha", ""))
                nArch = nArch + 1
                nHueco = 0
                ' Se vuelca cada 20 archivos: si esto se corta a medio camino,
                ' lo ya leido queda guardado y la proxima corrida sigue de ahi.
                If nBuf > 40000 Or (nArch Mod 20) = 0 Then
                    VolcarPrecios wsPre, buf, nBuf
                    nBuf = 0
                End If
                DoEvents
            End If
        End If
    Next dFe

    Paso "Volcando " & Format$(nFilas + nNue, "#,##0") & " filas a BD_Precios"
    VolcarPrecios wsPre, buf, nBuf

    Restaurar calcPrev, scrPrev, evPrev
    Application.DisplayAlerts = alertPrev

    t = Timer - mT0
    msg = "v" & VER & " - CargarVectores terminado en " & Format$(t, "0.0") & " s"
    If t > 90 Then msg = msg & " (" & Format$(t / 60, "0.0") & " min)"
    msg = msg & "." & vbCrLf & vbCrLf & mPasos & vbCrLf
    msg = msg & "Archivos leidos ahora: " & nArch & vbCrLf
    msg = msg & "Fechas que ya estaban: " & nYa & vbCrLf
    msg = msg & "Fechas sin archivo (fines de semana, feriados): " & nSalt & vbCrLf
    msg = msg & "Filas agregadas: " & Format$(nFilas, "#,##0") & vbCrLf
    msg = msg & "Codigos en el Maestro: " & codigos.Count & vbCrLf
    If nuevos.Count > 0 Then
        msg = msg & "Codigos nuevos (sin historico): " & nuevos.Count & vbCrLf
        msg = msg & "Filas que les traje del historico: " & Format$(nNue, "#,##0") & vbCrLf
    End If
    msg = msg & "BD_Precios queda con " & Format$(FilasPrecios(), "#,##0") & " filas." & vbCrLf & vbCrLf
    If nArch = 0 Then
        If nSalt > 0 Then
            msg = msg & "No encontre ningun archivo para esas fechas." & vbCrLf
            msg = msg & "Revisa Config C2 y que se llamen 'AAAAMMDD RFL.xls'." & vbCrLf & vbCrLf
        End If
    End If
    msg = msg & "Ahora: 'Datos del vector' si agregaste codigos."
    Aviso msg, "CargarVectores"
    Exit Sub

Falla:
    ' no se tira lo ya leido
    On Error Resume Next
    VolcarPrecios wsPre, buf, nBuf
    On Error GoTo 0
    Application.DisplayAlerts = alertPrev
    Reventar Nothing, calcPrev, scrPrev, evPrev, "CargarVectores"
End Sub


Private Function LeerVector(ruta As String, dFe As Double, vMap As Object, codigos As Object, _
                            ByRef buf() As Variant, ByRef nBuf As Long, ByRef capa As Long, _
                            notaHueco As String) As Long
    Dim wb As Workbook, ws As Worksheet, v As Variant
    Dim i As Long, cod As String, n As Long

    On Error GoTo Salir
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    Set ws = wb.Worksheets(1)
    v = LeerHoja(ws, MaxCol(vMap))
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If Not IsArray(v) Then Exit Function

    For i = 1 To UBound(v, 1)
        cod = SinGuiones(Txt(v(i, vMap("codigo"))))
        If Len(cod) > 0 Then
            If codigos.Exists(cod) Then
                nBuf = nBuf + 1
                If nBuf > capa Then
                    capa = capa * 2
                    ReDim Preserve buf(1 To capa, 1 To PR_COLS)
                End If
                buf(nBuf, 1) = dFe
                buf(nBuf, 2) = "'" & cod
                buf(nBuf, 3) = Txt(v(i, vMap("nemonico")))
                buf(nBuf, 4) = Txt(v(i, vMap("moneda")))
                buf(nBuf, 5) = NumOVacio(v(i, vMap("precio")))
                buf(nBuf, 6) = NumOVacio(v(i, vMap("interes")))
                buf(nBuf, 7) = NumOVacio(v(i, vMap("cupon")))
                buf(nBuf, 8) = FechaSegura(v(i, vMap("vencimiento")))
                buf(nBuf, 9) = NumOVacio(v(i, vMap("tir")))
                buf(nBuf, 10) = notaHueco
                If Not EsNumero(v(i, vMap("precio"))) Then
                    buf(nBuf, 10) = MasFlag(buf(nBuf, 10), "REVISAR: sin precio limpio")
                End If
                n = n + 1
            End If
        End If
    Next i
    LeerVector = n
    Exit Function

Salir:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function


Private Sub PrepararPrecios(ws As Worksheet)
    Dim cab As Variant
    If Len(Txt(ws.Range("A1").Value)) > 0 Then Exit Sub
    cab = Array("Fecha", "Codigo SBS", "Nemonico", "Moneda", "Precio limpio %", _
                "Interes acum", "Tasa cupon", "Vencimiento", "TIR %", "Flag")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PR_COLS)).Value = cab
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, PR_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
    End With
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 16
    ws.Columns(10).ColumnWidth = 42
    CongelarPaneles ws, 1, 2
End Sub


Private Function FechasCargadas(ws As Worksheet, ByRef ultFecha As Double) As Object
    Dim d As Object, ult As Long, v As Variant, i As Long
    Set d = NuevoDic()
    ultFecha = 0
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then
        Set FechasCargadas = d
        Exit Function
    End If
    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 1)).Value2
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If EsNumero(v(i, 1)) Then
                d(CLng(v(i, 1))) = 1
                If CDbl(v(i, 1)) > ultFecha Then ultFecha = CDbl(v(i, 1))
            End If
        Next i
    Else
        If EsNumero(v) Then
            d(CLng(v)) = 1
            ultFecha = CDbl(v)
        End If
    End If
    Set FechasCargadas = d
End Function


Private Sub VolcarPrecios(ws As Worksheet, buf() As Variant, nBuf As Long)
    Dim ult As Long, i As Long, j As Long, out() As Variant
    If nBuf = 0 Then Exit Sub

    ReDim out(1 To nBuf, 1 To PR_COLS)
    For i = 1 To nBuf
        For j = 1 To PR_COLS
            out(i, j) = buf(i, j)
        Next j
    Next i

    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 1 Then ult = 1
    ws.Range(ws.Cells(ult + 1, 1), ws.Cells(ult + nBuf, PR_COLS)).Value = out
    ' solo se formatea el tramo recien pegado, no toda la hoja
    ws.Range(ws.Cells(ult + 1, 1), ws.Cells(ult + nBuf, 1)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(ult + 1, 8), ws.Cells(ult + nBuf, 8)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(ult + 1, 5), ws.Cells(ult + nBuf, 7)).NumberFormat = "0.000000"
    ws.Range(ws.Cells(ult + 1, 9), ws.Cells(ult + nBuf, 9)).NumberFormat = "0.0000"
End Sub


Private Function FilasPrecios() As Long
    Dim ws As Worksheet, u As Long
    If Not HojaExiste(SH_PRE) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_PRE)
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If u > 1 Then FilasPrecios = u - 1
End Function


Private Function CodigosNuevos(codigos As Object) As Object
    ' Codigos del Maestro que no tienen ni una fila en BD_Precios.
    Dim d As Object, hay As Object, ws As Worksheet
    Dim ult As Long, i As Long, v As Variant, cod As String, k As Variant

    Set d = NuevoDic()
    Set CodigosNuevos = d
    If codigos Is Nothing Then Exit Function

    Set hay = NuevoDic()
    If HojaExiste(SH_PRE) Then
        Set ws = ThisWorkbook.Worksheets(SH_PRE)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ult >= 2 Then
            v = ws.Range(ws.Cells(2, 2), ws.Cells(ult, 2)).Value2
            If IsArray(v) Then
                For i = 1 To UBound(v, 1)
                    cod = SinGuiones(Txt(v(i, 1)))
                    If Len(cod) > 0 Then hay(cod) = 1
                Next i
            Else
                cod = SinGuiones(Txt(v))
                If Len(cod) > 0 Then hay(cod) = 1
            End If
        End If
    End If

    For Each k In codigos.Keys
        If Not hay.Exists(CStr(k)) Then d(CStr(k)) = 1
    Next k
End Function


Public Sub BorrarPreciosDesde()
    ' Borra de BD_Precios todas las filas con fecha >= la que escribas.
    ' Para cuando corrigieron un vector y hay que releerlo: borras desde esa
    ' fecha y vuelves a correr Cargar vectores.
    Dim ws As Worksheet, s As String, d As Double, ult As Long, i As Long, n As Long
    Dim v As Variant, keep() As Variant, k As Long, j As Long

    On Error GoTo Falla
    If Not HojaExiste(SH_PRE) Then
        MsgBox "No existe BD_Precios.", vbInformation, "BorrarPreciosDesde"
        Exit Sub
    End If
    s = InputBox("Borrar BD_Precios DESDE que fecha (dd/mm/aaaa)?" & vbCrLf & vbCrLf & _
                 "Se borran esa fecha y todas las posteriores." & vbCrLf & _
                 "Despues corre 'Cargar vectores' para releerlas.", "BorrarPreciosDesde")
    If Len(Trim$(s)) = 0 Then Exit Sub
    If Not IsDate(s) Then
        MsgBox "No entendi la fecha: " & s, vbExclamation, "BorrarPreciosDesde"
        Exit Sub
    End If
    d = CDbl(CDate(s))

    Set ws = ThisWorkbook.Worksheets(SH_PRE)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Sub
    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, PR_COLS)).Value
    ReDim keep(1 To UBound(v, 1), 1 To PR_COLS)
    k = 0
    For i = 1 To UBound(v, 1)
        If EsNumero(v(i, 1)) And CDbl(v(i, 1)) >= d Then
            n = n + 1
        Else
            k = k + 1
            For j = 1 To PR_COLS
                keep(k, j) = v(i, j)
            Next j
        End If
    Next i
    If n = 0 Then
        MsgBox "No hay filas con fecha >= " & Format$(d, "dd/mm/yyyy") & ".", vbInformation, "BorrarPreciosDesde"
        Exit Sub
    End If
    If MsgBox("Se van a borrar " & Format$(n, "#,##0") & " filas desde el " & _
              Format$(d, "dd/mm/yyyy") & ". Seguro?", vbYesNo + vbQuestion, "BorrarPreciosDesde") <> vbYes Then Exit Sub

    ws.Range(ws.Cells(2, 1), ws.Cells(ult, PR_COLS)).ClearContents
    If k > 0 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(k + 1, PR_COLS)).Value = keep
        ws.Range(ws.Cells(2, 1), ws.Cells(k + 1, 1)).NumberFormat = "dd/mm/yyyy"
        ws.Range(ws.Cells(2, 8), ws.Cells(k + 1, 8)).NumberFormat = "dd/mm/yyyy"
    End If
    MsgBox "Borradas " & Format$(n, "#,##0") & " filas. BD_Precios queda con " & _
           Format$(k, "#,##0") & "." & vbCrLf & vbCrLf & _
           "Ahora corre 'Cargar vectores'.", vbInformation, "BorrarPreciosDesde"
    Exit Sub

Falla:
    MsgBox "BorrarPreciosDesde se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "BorrarPreciosDesde"
End Sub


'==============================================================================
'  DATOS DEL VECTOR  ->  Maestro E:N
'==============================================================================
Public Sub CargarDatosVector()
    ' Lee el ultimo vector y llena en el Maestro las columnas E:N de cada
    ' codigo: nemonico, emisor, cupon, vencimiento, moneda, ISIN, rating,
    ' duracion, spread y la fecha en que lo vio.
    ' REGLA: el vector pisa solo lo que trae. Lo que no trae no se toca.
    Dim wsCfg As Worksheet, ws As Worksheet, wsV As Worksheet, wb As Workbook
    Dim carpeta As String, ruta As String, vMap As Object
    Dim dFe As Double, i As Long, ult As Long, fila As Long
    Dim v As Variant, idx As Object
    Dim cod As String, facPb As Double
    Dim nOk As Long, nSin As Long, msg As String
    Dim scrPrev As Boolean, calcPrev As XlCalculation, evPrev As Boolean

    mT0 = Timer: mPasos = "": mDetalle = ""
    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Paso "Buscando el ultimo vector"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    carpeta = SinBarra(Txt(wsCfg.Range("C2").Value))
    If Len(carpeta) = 0 Then Err.Raise vbObjectError + 60, , "Falta la carpeta del vector en Config C2."
    If Not HojaExiste(SH_MAE) Then Err.Raise vbObjectError + 63, , "No existe la hoja Maestro. Corre Crear Maestro."
    Set vMap = MapaVector(wsCfg)

    If IsDate(wsCfg.Range("C4").Value) Then
        dFe = VectorHastaFecha(carpeta, CDbl(CDate(wsCfg.Range("C4").Value)), ruta)
    Else
        dFe = UltimoVector(carpeta)
        If dFe > 0 Then ruta = RutaVector(carpeta, dFe)
    End If
    If Len(ruta) = 0 Then Err.Raise vbObjectError + 61, , "No encontre ningun vector en la carpeta."

    Paso "Leyendo " & Format$(CDate(dFe), "dd/mm/yyyy")
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    Set wsV = wb.Worksheets(1)
    v = LeerHoja(wsV, MaxCol(vMap))
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Set wsV = Nothing
    If Not IsArray(v) Then Err.Raise vbObjectError + 62, , "El vector salio vacio."

    ' codigo -> fila del vector (la primera que aparezca)
    Set idx = NuevoDic()
    For i = 1 To UBound(v, 1)
        cod = SinGuiones(Txt(v(i, vMap("codigo"))))
        If Len(cod) > 0 Then
            If Not idx.Exists(cod) Then idx(cod) = i
        End If
    Next i

    facPb = NumDef(wsCfg.Range("C21").Value, 100)
    If facPb = 0 Then facPb = 1

    Paso "Escribiendo el Maestro"
    Set ws = ThisWorkbook.Worksheets(SH_MAE)
    CabeceraMaestro ws
    ult = ws.Cells(ws.Rows.Count, MA_COD).End(xlUp).Row
    For fila = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(fila, MA_COD).Value))
        If Len(cod) > 0 Then
            If idx.Exists(cod) Then
                i = CLng(idx(cod))
                nOk = nOk + 1
                PonTxt ws, fila, MA_NEM, v(i, vMap("nemonico"))
                PonTxt ws, fila, MA_EMI, v(i, vMap("emisor"))
                PonNum ws, fila, MA_CUP, v(i, vMap("cupon")), 1
                PonFecha ws, fila, MA_VEN, v(i, vMap("vencimiento"))
                PonTxt ws, fila, MA_MON, NormMoneda(Txt(v(i, vMap("moneda"))))
                If vMap("isin") > 0 Then PonTxt ws, fila, MA_ISI, v(i, vMap("isin"))
                PonTxt ws, fila, MA_RAT, v(i, vMap("rating"))
                PonNum ws, fila, MA_DUR, v(i, vMap("duracion")), 1
                If vMap("spread") > 0 Then PonNum ws, fila, MA_SPR, v(i, vMap("spread")), facPb
                ws.Cells(fila, MA_VIS).Value = dFe
                ws.Cells(fila, MA_VIS).NumberFormat = "dd/mm/yyyy"
            Else
                nSin = nSin + 1
                ws.Cells(fila, MA_VIS).Value = "NO"
            End If
        End If
    Next fila

    Restaurar calcPrev, scrPrev, evPrev

    msg = "v" & VER & " - Datos del vector listos en " & Format$(Timer - mT0, "0.0") & " s."
    msg = msg & vbCrLf & vbCrLf & mPasos & vbCrLf
    msg = msg & "Vector leido: " & Format$(CDate(dFe), "dd/mm/yyyy") & vbCrLf
    msg = msg & "Codigos encontrados: " & nOk & vbCrLf
    If nSin > 0 Then
        msg = msg & "Codigos que NO estan en el vector: " & nSin & "  (columna N = NO)" & vbCrLf
        msg = msg & "Revisa que el codigo este bien escrito." & vbCrLf
    End If
    If vMap("isin") = 0 Then msg = msg & vbCrLf & "ISIN no se lleno: falta la letra en Config C19."
    If vMap("spread") = 0 Then msg = msg & vbCrLf & "Spread no se lleno: falta la letra en Config C20."
    Aviso msg, "CargarDatosVector"
    Exit Sub

Falla:
    Reventar wb, calcPrev, scrPrev, evPrev, "CargarDatosVector"
End Sub


Private Sub PonTxt(ws As Worksheet, fila As Long, col As Long, valor As Variant)
    Dim s As String
    s = Trim$(Txt(valor))
    If Len(s) > 0 Then ws.Cells(fila, col).Value = s
End Sub


Private Sub PonNum(ws As Worksheet, fila As Long, col As Long, valor As Variant, factor As Double)
    If EsNumero(valor) Then ws.Cells(fila, col).Value = CDbl(valor) * factor
End Sub


Private Sub PonFecha(ws As Worksheet, fila As Long, col As Long, valor As Variant)
    Dim f As Variant
    f = FechaSegura(valor)
    If Not IsEmpty(f) Then ws.Cells(fila, col).Value = f
End Sub


'==============================================================================
'  VER COLUMNAS DEL VECTOR
'==============================================================================
Public Sub ColumnasVector()
    ' Abre el ultimo vector y escribe, en la hoja Columnas_Vector, una fila por
    ' columna: letra, cabecera y tres valores de ejemplo. De ahi sacas las
    ' letras que van en Config C8:C20.
    Dim wsCfg As Worksheet, ws As Worksheet, wsV As Worksheet, wb As Workbook
    Dim carpeta As String, ruta As String, dFe As Double
    Dim i As Long, j As Long, ultF As Long, ultC As Long, fila As Long
    Dim v As Variant, cabFila As Long
    Dim scrPrev As Boolean, calcPrev As XlCalculation, evPrev As Boolean

    mT0 = Timer: mPasos = "": mDetalle = ""
    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Paso "Buscando el ultimo vector"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    carpeta = SinBarra(Txt(wsCfg.Range("C2").Value))
    If Len(carpeta) = 0 Then Err.Raise vbObjectError + 70, , "Falta la carpeta del vector en Config C2."
    dFe = UltimoVector(carpeta)
    If dFe > 0 Then ruta = RutaVector(carpeta, dFe)
    If Len(ruta) = 0 Then Err.Raise vbObjectError + 71, , "No encontre ningun vector en la carpeta."

    Paso "Leyendo " & Format$(CDate(dFe), "dd/mm/yyyy")
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    Set wsV = wb.Worksheets(1)
    ultF = wsV.UsedRange.Row + wsV.UsedRange.Rows.Count - 1
    ultC = wsV.UsedRange.Column + wsV.UsedRange.Columns.Count - 1
    If ultF > 12 Then ultF = 12
    If ultC > 120 Then ultC = 120
    v = wsV.Range(wsV.Cells(1, 1), wsV.Cells(ultF, ultC)).Value2
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Set wsV = Nothing
    If Not IsArray(v) Then Err.Raise vbObjectError + 72, , "El vector salio vacio."

    ' La cabecera es la primera fila con texto en al menos tres columnas.
    cabFila = 1
    For i = 1 To UBound(v, 1)
        j = 0
        For fila = 1 To UBound(v, 2)
            If Len(Txt(v(i, fila))) > 0 Then j = j + 1
        Next fila
        If j >= 3 Then
            cabFila = i
            Exit For
        End If
    Next i

    Paso "Escribiendo " & SH_CVE
    Set ws = HojaOCrea(SH_CVE)
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("Letra", "Num", "Cabecera", "Ejemplo 1", "Ejemplo 2", "Ejemplo 3")
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
    End With

    fila = 1
    For j = 1 To UBound(v, 2)
        fila = fila + 1
        ws.Cells(fila, 1).Value = NumALetra(j)
        ws.Cells(fila, 2).Value = j
        ws.Cells(fila, 3).Value = Txt(v(cabFila, j))
        For i = 1 To 3
            If cabFila + i <= UBound(v, 1) Then
                ws.Cells(fila, 3 + i).Value = Txt(v(cabFila + i, j))
            End If
        Next i
    Next j

    ws.Columns(1).ColumnWidth = 7
    ws.Columns(2).ColumnWidth = 6
    ws.Columns(3).ColumnWidth = 40
    ws.Range(ws.Columns(4), ws.Columns(6)).ColumnWidth = 22
    CongelarPaneles ws, 1, 3

    ws.Range("H1").Value = "Archivo: " & ruta
    ws.Range("H2").Value = "Cabecera leida en la fila " & cabFila & "."
    ws.Range("H3").Value = "Busca ISIN, duracion, spread y rating, y copia la letra a Config C8:C20."
    ws.Range("H1:H3").Font.Color = RGB(100, 100, 100)
    ws.Columns(8).ColumnWidth = 70

    Restaurar calcPrev, scrPrev, evPrev
    ws.Activate

    Aviso "v" & VER & " - ColumnasVector listo." & vbCrLf & vbCrLf & _
          "Vector leido: " & Format$(CDate(dFe), "dd/mm/yyyy") & vbCrLf & _
          "Columnas: " & UBound(v, 2) & vbCrLf & vbCrLf & _
          "Mira la hoja " & SH_CVE & " y copia las letras a Config.", "ColumnasVector"
    Exit Sub

Falla:
    Reventar wb, calcPrev, scrPrev, evPrev, "ColumnasVector"
End Sub


'==============================================================================
'  DIAGNOSTICO
'==============================================================================
Public Sub Diagnostico()
    Dim wsCfg As Worksheet, ws As Worksheet, msg As String
    Dim codigos As Object, nuevos As Object, k As Variant
    Dim ult As Long, i As Long, v As Variant
    Dim dMin As Double, dMax As Double, nFechas As Object, nSinPrecio As Long
    Dim carpeta As String, dUlt As Double, nNo As Long

    On Error GoTo Falla
    msg = "BASE RF  v" & VER & vbCrLf & String$(44, "-") & vbCrLf

    If Not HojaExiste(SH_CFG) Then
        msg = msg & "No existe Config. Corre Crear Config." & vbCrLf
        MsgBox msg, vbExclamation, "Diagnostico"
        Exit Sub
    End If
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    carpeta = SinBarra(Txt(wsCfg.Range("C2").Value))
    msg = msg & "Carpeta del vector: " & IIf(Len(carpeta) > 0, carpeta, "VACIA (Config C2)") & vbCrLf
    If Len(carpeta) > 0 Then
        dUlt = UltimoVector(carpeta)
        msg = msg & "Ultimo vector en la carpeta: " & _
              IIf(dUlt > 0, Format$(CDate(dUlt), "dd/mm/yyyy"), "NINGUNO - revisa la carpeta") & vbCrLf
    End If
    msg = msg & "Inicio del historico (C3): " & Txt(wsCfg.Range("C3").Text) & vbCrLf
    msg = msg & "ISIN: col " & IIf(Len(Txt(wsCfg.Range("C19").Value)) > 0, Txt(wsCfg.Range("C19").Value), "SIN LETRA") & _
          "   Spread: col " & IIf(Len(Txt(wsCfg.Range("C20").Value)) > 0, Txt(wsCfg.Range("C20").Value), "SIN LETRA") & vbCrLf & vbCrLf

    Set codigos = CodigosMaestro()
    msg = msg & "MAESTRO: " & codigos.Count & " codigos" & vbCrLf
    If HojaExiste(SH_MAE) Then
        Set ws = ThisWorkbook.Worksheets(SH_MAE)
        For Each k In codigos.Keys
            If UCase$(Txt(ws.Cells(CLng(codigos(k)), MA_VIS).Value)) = "NO" Then nNo = nNo + 1
        Next k
        If nNo > 0 Then msg = msg & "   " & nNo & " no estan en el vector (columna N = NO)" & vbCrLf
    End If
    Set nuevos = CodigosNuevos(codigos)
    If nuevos.Count > 0 Then
        msg = msg & "   " & nuevos.Count & " sin ninguna fila en BD_Precios -> corre Cargar vectores" & vbCrLf
    End If
    msg = msg & vbCrLf

    If HojaExiste(SH_PRE) Then
        Set ws = ThisWorkbook.Worksheets(SH_PRE)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        msg = msg & "BD_PRECIOS: " & Format$(IIf(ult > 1, ult - 1, 0), "#,##0") & " filas" & vbCrLf
        If ult > 1 Then
            Set nFechas = NuevoDic()
            v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 5)).Value2
            For i = 1 To UBound(v, 1)
                If EsNumero(v(i, 1)) Then
                    nFechas(CLng(v(i, 1))) = 1
                    If dMin = 0 Or CDbl(v(i, 1)) < dMin Then dMin = CDbl(v(i, 1))
                    If CDbl(v(i, 1)) > dMax Then dMax = CDbl(v(i, 1))
                End If
                If Not EsNumero(v(i, 5)) Then nSinPrecio = nSinPrecio + 1
            Next i
            msg = msg & "   del " & Format$(CDate(dMin), "dd/mm/yyyy") & " al " & _
                  Format$(CDate(dMax), "dd/mm/yyyy") & "  (" & nFechas.Count & " fechas)" & vbCrLf
            If nSinPrecio > 0 Then msg = msg & "   " & nSinPrecio & " filas sin precio limpio (Flag)" & vbCrLf
            If dUlt > dMax Then
                msg = msg & "   FALTA cargar: hay vectores hasta el " & Format$(CDate(dUlt), "dd/mm/yyyy") & vbCrLf
            End If
        End If
    Else
        msg = msg & "BD_PRECIOS: no existe todavia. Corre Cargar vectores." & vbCrLf
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

    ws.Range("B2").Value = "BASE RF  -  vector SBS"
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
        "CargarVectores", "1. Cargar vectores", _
        "CargarDatosVector", "2. Datos del vector")

    extra = Array( _
        "NormalizarExcel", "ARREGLAR EXCEL", _
        "CrearConfig", "Crear Config", _
        "CrearMaestro", "Crear Maestro", _
        "ColumnasVector", "Ver columnas del vector", _
        "RefrescarIndice", "Refrescar indice", _
        "BorrarPreciosDesde", "Borrar precios desde...", _
        "Diagnostico", "Diagnostico", _
        "Version", "Version")

    BotonesEnColumna ws, rutina, 60, 110, 170, 30, True
    BotonesEnColumna ws, extra, 300, 110, 170, 26, False

    ws.Range("B26").Value = "1 todos los dias. 2 solo cuando agregas un codigo al Maestro o cambias una letra en Config."
    ws.Range("B27").Value = "Si el 1 dice que no hay nada nuevo y si copiaste vectores: Refrescar indice."
    ws.Range("B28").Value = "Si corrigieron un vector: Borrar precios desde esa fecha y luego el 1."
    ws.Range("B29").Value = "Si cortaste una macro con Esc: ARREGLAR EXCEL antes de seguir."
    ws.Range("B26:B29").Font.Color = RGB(128, 128, 128)

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
    ' Si una macro se corta a medias (Esc + Finalizar), Excel se queda en
    ' calculo MANUAL para TODOS los libros. Esto lo devuelve a la normalidad.
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
    MsgBox "mod_BaseRF  v" & VER, vbInformation, "Version"
End Sub


Public Sub RefrescarIndice()
    Set mIdxVec = Nothing
    mIdxCarpeta = ""
    MsgBox "Indice de la carpeta borrado." & vbCrLf & _
           "La proxima corrida vuelve a leer la lista de archivos.", vbInformation, "RefrescarIndice"
End Sub


'==============================================================================
'  CARPETA DEL VECTOR
'==============================================================================
Private Function RutaVector(carpeta As String, d As Double) As String
    Dim idx As Object
    Set idx = IndiceVectores(carpeta)
    If idx.Exists(CLng(d)) Then RutaVector = CStr(idx(CLng(d)))
End Function


Private Function UltimoVector(carpeta As String) As Double
    ' La fecha mas alta que haya en la carpeta.
    Dim idx As Object, k As Variant
    Set idx = IndiceVectores(carpeta)
    For Each k In idx.Keys
        If CDbl(k) > UltimoVector Then UltimoVector = CDbl(k)
    Next k
End Function


Private Function VectorHastaFecha(carpeta As String, d As Double, ByRef ruta As String) As Double
    ' El ultimo vector con fecha <= d, buscando hasta 20 dias atras.
    Dim i As Long
    For i = 0 To 20
        ruta = RutaVector(carpeta, d - i)
        If Len(ruta) > 0 Then
            VectorHastaFecha = d - i
            Exit Function
        End If
    Next i
    ruta = ""
End Function


Private Function IndiceVectores(carpeta As String) As Object
    ' Lee el nombre de TODOS los vectores de la carpeta en una sola pasada y se
    ' lo guarda 10 minutos. Un solo golpe a la red por corrida.
    Dim d As Object, nom As String, ocho As String, dt As Double

    If Not mIdxVec Is Nothing Then
        If StrComp(mIdxCarpeta, carpeta, vbTextCompare) = 0 Then
            If Timer >= mIdxStamp And Timer - mIdxStamp < 600 Then
                Set IndiceVectores = mIdxVec
                Exit Function
            End If
        End If
    End If

    Set d = NuevoDic()
    mDetalle = "Indexando la carpeta del vector"
    nom = Dir(carpeta & Application.PathSeparator & "*RFL*.xls*")
    Do While Len(nom) > 0
        ocho = OchoDigitos(nom)
        If Len(ocho) = 8 Then
            dt = 0
            On Error Resume Next
            dt = CDbl(DateSerial(CLng(Left$(ocho, 4)), CLng(Mid$(ocho, 5, 2)), CLng(Right$(ocho, 2))))
            On Error GoTo 0
            If dt > 0 Then
                If Not d.Exists(CLng(dt)) Then d(CLng(dt)) = carpeta & Application.PathSeparator & nom
            End If
        End If
        nom = Dir
    Loop

    Set mIdxVec = d
    mIdxCarpeta = carpeta
    mIdxStamp = Timer
    Set IndiceVectores = d
End Function


Private Function OchoDigitos(ByVal nom As String) As String
    Dim i As Long, c As String, run As String
    For i = 1 To Len(nom)
        c = Mid$(nom, i, 1)
        If c >= "0" And c <= "9" Then
            run = run & c
            If Len(run) = 8 Then
                OchoDigitos = run
                Exit Function
            End If
        Else
            run = ""
        End If
    Next i
End Function


Private Function MapaVector(wsCfg As Worksheet) As Object
    Dim d As Object
    Set d = NuevoDic()
    d("codigo") = Letra(Txt(wsCfg.Range("C8").Value), 1)
    d("nemonico") = Letra(Txt(wsCfg.Range("C9").Value), 3)
    d("emisor") = Letra(Txt(wsCfg.Range("C10").Value), 5)
    d("moneda") = Letra(Txt(wsCfg.Range("C11").Value), 6)
    d("precio") = Letra(Txt(wsCfg.Range("C12").Value), 10)
    d("interes") = Letra(Txt(wsCfg.Range("C13").Value), 13)
    d("tir") = Letra(Txt(wsCfg.Range("C14").Value), 14)
    d("vencimiento") = Letra(Txt(wsCfg.Range("C15").Value), 17)
    d("cupon") = Letra(Txt(wsCfg.Range("C16").Value), 18)
    d("duracion") = Letra(Txt(wsCfg.Range("C17").Value), 24)
    d("rating") = Letra(Txt(wsCfg.Range("C18").Value), 21)
    d("isin") = Letra(Txt(wsCfg.Range("C19").Value), 0)
    d("spread") = Letra(Txt(wsCfg.Range("C20").Value), 0)
    Set MapaVector = d
End Function


Private Function MaxCol(m As Object) As Long
    Dim k As Variant
    If m Is Nothing Then Exit Function
    For Each k In m.Keys
        If CLng(m(k)) > MaxCol Then MaxCol = CLng(m(k))
    Next k
End Function


Private Function LeerHoja(ws As Worksheet, maxC As Long) As Variant
    Dim ult As Long
    If maxC < 1 Then maxC = 1
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ult < 1 Then Exit Function
    LeerHoja = ws.Range(ws.Cells(1, 1), ws.Cells(ult, maxC)).Value2
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
    Application.StatusBar = "Base RF: " & texto & "..."
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


Private Function SinBarra(ByVal s As String) As String
    s = Trim$(s)
    Do While Len(s) > 0 And Right$(s, 1) = Application.PathSeparator
        s = Left$(s, Len(s) - 1)
    Loop
    SinBarra = s
End Function


Private Function NormMoneda(ByVal t As String) As String
    t = UCase$(Trim$(t))
    If InStr(t, "USD") > 0 Or InStr(t, "US$") > 0 Or InStr(t, "DOL") > 0 Then
        NormMoneda = "USD"
    ElseIf InStr(t, "PEN") > 0 Or InStr(t, "SOL") > 0 Or InStr(t, "S/") > 0 Then
        NormMoneda = "PEN"
    ElseIf InStr(t, "VAC") > 0 Then
        NormMoneda = "VAC"
    Else
        NormMoneda = t
    End If
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


Private Function NumOVacio(v As Variant) As Variant
    If EsNumero(v) Then
        NumOVacio = CDbl(v)
    Else
        NumOVacio = Empty
    End If
End Function


Private Function FechaSegura(v As Variant) As Variant
    ' el vector trae fecha de Excel; por si acaso tambien AAAAMMDD
    Dim s As String
    If EsNumero(v) Then
        If CDbl(v) > 19000101 And CDbl(v) < 21001231 Then
            s = Format$(CDbl(v), "00000000")
            FechaSegura = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Right$(s, 2)))
        Else
            FechaSegura = CDbl(v)
        End If
        Exit Function
    End If
    s = Trim$(Txt(v))
    If Len(s) = 8 And IsNumeric(s) Then
        FechaSegura = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Right$(s, 2)))
    ElseIf IsDate(s) Then
        FechaSegura = CDate(s)
    Else
        FechaSegura = Empty
    End If
End Function


Private Function MasFlag(actual As Variant, nuevo As String) As String
    Dim a As String
    a = Txt(actual)
    If Len(a) = 0 Then
        MasFlag = nuevo
    Else
        MasFlag = a & " | " & nuevo
    End If
End Function
