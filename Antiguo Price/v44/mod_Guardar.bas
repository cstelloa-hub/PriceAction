Attribute VB_Name = "mod_Guardar"
Option Explicit

'==============================================================================
' mod_Guardar  v2.1  (para mod_Iliquidos v2.70+, un solo libro)
'
' Modulo aparte. No llama a nada de mod_Iliquidos: se puede pegar y borrar
' sin romper el otro.
'
' Que hace (boton "Guardar la foto"):
'   Crea un .xlsx nuevo con tres hojas, listo para pegar en el otro Excel:
'     CUADRO         el cuadro tal cual. Las filas grises (subtotales, totales,
'                    niveles) conservan sus formulas SUMPRODUCT, que apuntan a
'                    celdas visibles de la misma hoja. Las filas de cada bono o
'                    fondo van en valores (su formula apuntaba a Retornos /
'                    Fondos, que no viajan).
'     Matriz_RF      precio limpio por codigo (filas) y fecha (columnas), con
'                    moneda y tasa cupon. Sale de BD_Precios. Solo los codigos
'                    que estan en el cuadro, desde la base del FY hasta el corte.
'     Matriz_Fondos  var ajustada por fondo (filas) y fecha de marca (columnas).
'                    Sale de Marcas_Detalle.
'   Sin nombres definidos ni vinculos al libro original.
'
' Donde guarda:
'   Config C14 si tiene una carpeta escrita; si no, <carpeta del libro>\Resumenes.
'   Nombre: Iliquidos_CUADRO_<corte yyyymmdd>.xlsx  (_v2, _v3 si ya existe)
'
' Corte: Config C4; vacio = ultima fecha de BD_Precios.
' Inicio de las matrices: Config C7 (base del FY); vacio = 31/10 del ano anterior.
'==============================================================================

Private Const G_CFG As String = "Config"
Private Const G_CUA As String = "CUADRO"
Private Const G_PRE As String = "BD_Precios"
Private Const G_MDE As String = "Marcas_Detalle"
Private Const G_INS As String = "Instrumentos"
Private Const G_MRF As String = "Matriz_RF"
Private Const G_MFO As String = "Matriz_Fondos"

Private mPaso As String


Public Sub GuardarCuadro()
    Dim wbNew As Workbook, wsOri As Worksheet, wsNew As Worksheet
    Dim carpeta As String, base As String, ruta As String
    Dim dCorte As Date, dIni As Date, n As Long
    Dim scrPrev As Boolean, evPrev As Boolean, alPrev As Boolean, calcPrev As XlCalculation
    Dim t0 As Single, msg As String, nF As Long, nV As Long, nRF As Long, nFo As Long

    On Error GoTo Falla
    t0 = Timer
    mPaso = "arrancando"
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    alPrev = Application.DisplayAlerts
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If Not HojaExiste(G_CUA) Then Err.Raise vbObjectError + 101, , _
        "No encuentro la hoja '" & G_CUA & "'. Corre RECALCULAR primero."
    Set wsOri = ThisWorkbook.Worksheets(G_CUA)
    If wsOri.Visible <> xlSheetVisible Then wsOri.Visible = xlSheetVisible

    mPaso = "leyendo corte e inicio de Config"
    dCorte = CorteDeConfig()
    dIni = InicioDeConfig(dCorte)

    ' ---- carpeta y nombre
    mPaso = "preparando la carpeta destino"
    carpeta = SinBarraFinal(Texto(ThisWorkbook.Worksheets(G_CFG).Range("C14").Value))
    If Len(carpeta) = 0 Then
        If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 102, , _
            "El libro no esta guardado y Config C14 esta vacio: no se donde guardar la foto."
        carpeta = ThisWorkbook.Path & Application.PathSeparator & "Resumenes"
    End If
    If Len(Dir(carpeta, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir carpeta
        On Error GoTo Falla
        If Len(Dir(carpeta, vbDirectory)) = 0 Then Err.Raise vbObjectError + 103, , _
            "No existe la carpeta y no pude crearla: " & carpeta & vbCrLf & "Revisa Config C14."
    End If
    base = carpeta & Application.PathSeparator & "Iliquidos_CUADRO_" & Format$(dCorte, "yyyymmdd")
    ruta = base & ".xlsx"
    n = 1
    Do While Len(Dir(ruta)) > 0
        n = n + 1
        ruta = base & "_v" & n & ".xlsx"
    Loop

    ' ---- 1. copia del CUADRO (formato, anchos, colores) a un libro nuevo
    mPaso = "copiando la hoja CUADRO a un libro nuevo"
    wsOri.Copy
    Set wbNew = ActiveWorkbook
    ' Seguro: si por lo que sea el activo sigue siendo este libro, no se toca.
    If wbNew Is ThisWorkbook Then
        Set wbNew = Nothing
        Err.Raise vbObjectError + 104, , "La copia de la hoja no creo un libro nuevo."
    End If
    Set wsNew = wbNew.Worksheets(1)
    Application.Calculation = xlCalculationManual

    ' ---- 2. formulas: las que solo miran la misma hoja se quedan; las que
    '         miran otra hoja o un nombre definido pasan a valor.
    mPaso = "pasando las formulas del CUADRO"
    CongelarFormulas wsOri, wsNew, nF, nV
    mPaso = "borrando nombres definidos de la copia"
    BorrarNombres wbNew

    ' ---- 3. matrices
    mPaso = "armando " & G_MRF
    nRF = MatrizRF(wbNew, dIni, dCorte)
    mPaso = "armando " & G_MFO
    nFo = MatrizFondos(wbNew, dIni, dCorte)

    mPaso = "guardando " & ruta
    On Error Resume Next
    wsNew.Activate
    wsNew.Range("A1").Select
    On Error GoTo Falla
    Application.Calculation = xlCalculationAutomatic
    wbNew.Calculate

    Application.DisplayAlerts = False
    wbNew.SaveAs Filename:=ruta, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

    Application.Calculation = calcPrev
    Application.DisplayAlerts = alPrev
    Application.EnableEvents = evPrev
    Application.ScreenUpdating = scrPrev

    msg = "Foto guardada en " & Format$(Timer - t0, "0.0") & " s." & vbCrLf & vbCrLf
    msg = msg & ruta & vbCrLf & vbCrLf
    msg = msg & "Corte: " & Format$(dCorte, "dd/mm/yyyy") & "   Matrices desde: " & Format$(dIni, "dd/mm/yyyy") & vbCrLf
    msg = msg & "CUADRO: " & nF & " formulas vivas (filas grises), " & nV & " celdas pasadas a valor." & vbCrLf
    msg = msg & G_MRF & ": " & nRF & " codigos.   " & G_MFO & ": " & nFo & " fondos." & vbCrLf
    msg = msg & "Sin nombres definidos ni vinculos a este libro."
    MsgBox msg, vbInformation, "GuardarCuadro"
    Exit Sub

Falla:
    msg = "GuardarCuadro se detuvo en: " & mPaso & vbCrLf & vbCrLf & _
          "ESCRIBEME ESTA LINEA:   G" & Err.Number & " @ " & mPaso & vbCrLf & vbCrLf & Err.Description
    On Error Resume Next
    ' Nunca se cierra este libro: solo la copia, y solo si de verdad es otra.
    If Not wbNew Is Nothing Then
        If Not (wbNew Is ThisWorkbook) Then wbNew.Close SaveChanges:=False
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox msg, vbCritical, "GuardarCuadro"
End Sub


'------------------------------------------------------------- CUADRO: formulas

Private Sub CongelarFormulas(wsOri As Worksheet, wsNew As Worksheet, ByRef nF As Long, ByRef nV As Long)
    ' Recorre las formulas del CUADRO original. Para cada una:
    '   - reemplaza los nombres definidos de UNA celda (cfgPrimaPb, cfgSpread,
    '     peTC, peTotF1...) y las referencias a una celda de otra hoja
    '     (Config!$C$72, BD_Pesos!$M$3) por su valor de hoy;
    '   - si despues de eso la formula sigue mirando otra hoja o un nombre de
    '     rango (reSuma, foVent...), la celda pasa a valor;
    '   - si solo mira celdas de la misma hoja, se queda como formula.
    Dim rg As Range, c As Range, f As String, f2 As String
    Dim escalares As Object, rangos As Object, nm As Name, r As Range

    Set escalares = CreateObject("Scripting.Dictionary")
    Set rangos = CreateObject("Scripting.Dictionary")
    For Each nm In ThisWorkbook.Names
        Set r = Nothing
        On Error Resume Next
        Set r = nm.RefersToRange
        On Error GoTo 0
        If Not r Is Nothing Then
            If r.Cells.Count = 1 Then
                escalares(nm.Name) = r.Value
            Else
                rangos(nm.Name) = 1
            End If
        Else
            rangos(nm.Name) = 1
        End If
    Next nm

    On Error Resume Next
    Set rg = wsOri.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rg Is Nothing Then Exit Sub

    For Each c In rg.Cells
        f = c.Formula
        f2 = SustituirNombres(f, escalares)
        f2 = SustituirCeldasExternas(f2)
        If InStr(f2, "!") > 0 Or TieneNombreRango(f2, rangos) Then
            wsNew.Range(c.Address).Value = c.Value
            nV = nV + 1
        Else
            wsNew.Range(c.Address).Formula = f2
            nF = nF + 1
        End If
    Next c
End Sub


Private Function SustituirNombres(f As String, escalares As Object) As String
    ' Nombre de una celda -> su valor, solo si aparece como palabra completa.
    Dim k As Variant, s As String, p As Long, nombre As String, antes As String, despues As String
    s = f
    For Each k In escalares.Keys
        nombre = CStr(k)
        p = InStr(1, s, nombre, vbTextCompare)
        Do While p > 0
            antes = "": despues = ""
            If p > 1 Then antes = Mid$(s, p - 1, 1)
            If p + Len(nombre) <= Len(s) Then despues = Mid$(s, p + Len(nombre), 1)
            If Not EsLetraONumero(antes) And Not EsLetraONumero(despues) And despues <> "(" Then
                s = Left$(s, p - 1) & NumTexto(escalares(k)) & Mid$(s, p + Len(nombre))
                p = InStr(p + 1, s, nombre, vbTextCompare)
            Else
                p = InStr(p + Len(nombre), s, nombre, vbTextCompare)
            End If
        Loop
    Next k
    SustituirNombres = s
End Function


Private Function SustituirCeldasExternas(f As String) As String
    ' Hoja!$C$72 (una sola celda, no un rango) -> su valor de hoy.
    Dim re As Object, ms As Object, m As Object, s As String, hoja As String, celda As String
    Dim v As Variant, i As Long, esRango As Boolean
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "([A-Za-z_][A-Za-z0-9_]*|'[^']+')!\$?([A-Z]{1,3})\$?(\d{1,7})(?![:\d])"
    s = f
    Set ms = re.Execute(s)
    ' de atras hacia adelante para que las posiciones no se muevan
    For i = ms.Count - 1 To 0 Step -1
        Set m = ms(i)
        hoja = m.SubMatches(0)
        If Left$(hoja, 1) = "'" Then hoja = Mid$(hoja, 2, Len(hoja) - 2)
        celda = m.SubMatches(1) & m.SubMatches(2)
        ' si justo antes hay ":" es la segunda parte de un rango: no tocar
        esRango = False
        If m.FirstIndex > 0 Then
            If Mid$(s, m.FirstIndex, 1) = ":" Then esRango = True
        End If
        If Not esRango Then
            If HojaExiste(hoja) Then
                v = ThisWorkbook.Worksheets(hoja).Range(celda).Value
                s = Left$(s, m.FirstIndex) & NumTexto(v) & Mid$(s, m.FirstIndex + m.Length + 1)
            End If
        End If
    Next i
    SustituirCeldasExternas = s
End Function


Private Function TieneNombreRango(f As String, rangos As Object) As Boolean
    Dim k As Variant, p As Long, nombre As String, antes As String, despues As String
    For Each k In rangos.Keys
        nombre = CStr(k)
        p = InStr(1, f, nombre, vbTextCompare)
        Do While p > 0
            antes = "": despues = ""
            If p > 1 Then antes = Mid$(f, p - 1, 1)
            If p + Len(nombre) <= Len(f) Then despues = Mid$(f, p + Len(nombre), 1)
            If Not EsLetraONumero(antes) And Not EsLetraONumero(despues) Then
                TieneNombreRango = True
                Exit Function
            End If
            p = InStr(p + Len(nombre), f, nombre, vbTextCompare)
        Loop
    Next k
End Function


Private Function EsLetraONumero(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    EsLetraONumero = (ch Like "[A-Za-z0-9_.]")
End Function


Private Function NumTexto(v As Variant) As String
    ' Numero como lo entiende .Formula (punto decimal). Texto va entre comillas.
    If IsError(v) Then
        NumTexto = "0"
    ElseIf IsEmpty(v) Then
        NumTexto = "0"
    ElseIf VarType(v) = vbString Then
        If IsNumeric(v) Then
            NumTexto = Trim$(Str$(CDbl(v)))
        Else
            NumTexto = """" & Replace(CStr(v), """", """""") & """"
        End If
    ElseIf VarType(v) = vbBoolean Then
        NumTexto = IIf(v, "TRUE", "FALSE")
    Else
        NumTexto = Trim$(Str$(CDbl(v)))
    End If
    If Left$(NumTexto, 1) = "." Then NumTexto = "0" & NumTexto
    If Left$(NumTexto, 2) = "-." Then NumTexto = "-0" & Mid$(NumTexto, 2)
End Function


'------------------------------------------------------------------ matrices

Private Function CodigosDelCuadro(ByRef nombres As Object) As Object
    ' Codigos con categoria en Instrumentos, en el orden de la hoja.
    ' nombres: codigo -> nombre final (col J; si no, B; si no, E)
    Dim ws As Worksheet, i As Long, ult As Long, cod As String, nom As String, lst As Object
    Set lst = CreateObject("Scripting.Dictionary")
    lst.CompareMode = 1
    Set nombres = CreateObject("Scripting.Dictionary")
    nombres.CompareMode = 1
    Set CodigosDelCuadro = lst
    If Not HojaExiste(G_INS) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(G_INS)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuionesG(Texto(ws.Cells(i, 1).Value))
        If Len(cod) > 0 And Len(Texto(ws.Cells(i, 3).Value)) > 0 Then
            If Not lst.Exists(cod) Then
                lst(cod) = lst.Count + 1
                nom = Texto(ws.Cells(i, 10).Value)
                If Len(nom) = 0 Then nom = Texto(ws.Cells(i, 2).Value)
                If Len(nom) = 0 Then nom = Texto(ws.Cells(i, 5).Value)
                If Len(nom) = 0 Then nom = cod
                nombres(cod) = nom
            End If
        End If
    Next i
End Function


Private Function MatrizRF(wb As Workbook, dIni As Date, dCorte As Date) As Long
    ' BD_Precios: A fecha  B codigo  C nemonico  D moneda  E precio limpio
    '             F interes  G cupon  H vencimiento  I TIR  J flag
    Dim ws As Worksheet, v As Variant, ult As Long, i As Long
    Dim codigos As Object, nombres As Object, fechas As Object, datos As Object
    Dim cod As String, d As Double, k As String, mon As Object, cup As Object
    Dim fe() As Double, nFe As Long, j As Long, out() As Variant, r As Long
    Dim wsM As Worksheet, kk As Variant, lista() As String, nL As Long

    If Not HojaExiste(G_PRE) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(G_PRE)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function
    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 7)).Value2

    Set codigos = CodigosDelCuadro(nombres)
    Set fechas = CreateObject("Scripting.Dictionary")
    Set datos = CreateObject("Scripting.Dictionary"): datos.CompareMode = 1
    Set mon = CreateObject("Scripting.Dictionary"): mon.CompareMode = 1
    Set cup = CreateObject("Scripting.Dictionary"): cup.CompareMode = 1

    For i = 1 To UBound(v, 1)
        If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
            d = CDbl(v(i, 1))
            If d >= CDbl(dIni) And d <= CDbl(dCorte) Then
                cod = SinGuionesG(Texto(v(i, 2)))
                If codigos.Exists(cod) Then
                    If Not fechas.Exists(d) Then fechas(d) = 1
                    k = cod & "|" & CStr(d)
                    If IsNumeric(v(i, 5)) And Not IsEmpty(v(i, 5)) Then datos(k) = CDbl(v(i, 5))
                    If Not mon.Exists(cod) Then mon(cod) = Texto(v(i, 4))
                    If IsNumeric(v(i, 7)) And Not IsEmpty(v(i, 7)) Then cup(cod) = CDbl(v(i, 7))
                End If
            End If
        End If
    Next i
    If fechas.Count = 0 Or codigos.Count = 0 Then Exit Function

    nFe = fechas.Count
    ReDim fe(1 To nFe)
    j = 0
    For Each kk In fechas.Keys
        j = j + 1
        fe(j) = CDbl(kk)
    Next kk
    OrdenarDoubles fe

    ' solo los codigos que tienen al menos un precio
    ReDim lista(1 To codigos.Count)
    For Each kk In codigos.Keys
        If mon.Exists(CStr(kk)) Then
            nL = nL + 1
            lista(nL) = CStr(kk)
        End If
    Next kk
    If nL = 0 Then Exit Function

    ReDim out(1 To nL + 1, 1 To nFe + 4)
    out(1, 1) = "Codigo SBS": out(1, 2) = "Nombre": out(1, 3) = "Moneda": out(1, 4) = "Tasa cupon"
    For j = 1 To nFe
        out(1, 4 + j) = fe(j)
    Next j
    For r = 1 To nL
        cod = lista(r)
        out(r + 1, 1) = "'" & cod
        out(r + 1, 2) = nombres(cod)
        out(r + 1, 3) = mon(cod)
        If cup.Exists(cod) Then out(r + 1, 4) = cup(cod)
        For j = 1 To nFe
            k = cod & "|" & CStr(fe(j))
            If datos.Exists(k) Then out(r + 1, 4 + j) = datos(k)
        Next j
    Next r

    Set wsM = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsM.Name = G_MRF
    wsM.Range(wsM.Cells(1, 1), wsM.Cells(nL + 1, nFe + 4)).Value = out
    With wsM
        .Range(.Cells(1, 1), .Cells(1, nFe + 4)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, nFe + 4)).Interior.Color = RGB(64, 64, 64)
        .Range(.Cells(1, 1), .Cells(1, nFe + 4)).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(1, 5), .Cells(1, nFe + 4)).NumberFormat = "dd/mm/yy"
        .Range(.Cells(1, 5), .Cells(1, nFe + 4)).HorizontalAlignment = xlCenter
        .Range(.Cells(2, 4), .Cells(nL + 1, 4)).NumberFormat = "0.000"
        .Range(.Cells(2, 5), .Cells(nL + 1, nFe + 4)).NumberFormat = "0.0000"
        .Columns(1).ColumnWidth = 16
        .Columns(2).ColumnWidth = 34
        .Columns(3).ColumnWidth = 8
        .Columns(4).ColumnWidth = 9
        .Range(.Columns(5), .Columns(nFe + 4)).ColumnWidth = 9
        .Cells(nL + 3, 1).Value = "Precio limpio (BD_Precios col. E) por fecha. Retorno de una ventana = (P fin / P base - 1) + tasa cupon x dias / 365 / P base, como en la hoja Retornos."
        .Cells(nL + 3, 1).Font.Color = RGB(128, 128, 128)
    End With
    On Error Resume Next
    wsM.Activate
    ActiveWindow.FreezePanes = False
    wsM.Range("E2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    MatrizRF = nL
End Function


Private Function MatrizFondos(wb As Workbook, dIni As Date, dCorte As Date) As Long
    ' Marcas_Detalle: A codigo  B fecha  C var ajustada  D nombre en Marcas  E ln
    Dim ws As Worksheet, v As Variant, ult As Long, i As Long
    Dim codigos As Object, nombres As Object, fechas As Object, datos As Object, tiene As Object
    Dim cod As String, d As Double, k As String
    Dim fe() As Double, nFe As Long, j As Long, out() As Variant, r As Long
    Dim wsM As Worksheet, kk As Variant, lista() As String, nL As Long

    If Not HojaExiste(G_MDE) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(G_MDE)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function
    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 3)).Value2

    Set codigos = CodigosDelCuadro(nombres)
    Set fechas = CreateObject("Scripting.Dictionary")
    Set datos = CreateObject("Scripting.Dictionary"): datos.CompareMode = 1
    Set tiene = CreateObject("Scripting.Dictionary"): tiene.CompareMode = 1

    For i = 1 To UBound(v, 1)
        cod = SinGuionesG(Texto(v(i, 1)))
        If Len(cod) > 0 And IsNumeric(v(i, 2)) And Not IsEmpty(v(i, 2)) Then
            d = CDbl(v(i, 2))
            If d >= CDbl(dIni) And d <= CDbl(dCorte) And codigos.Exists(cod) Then
                If IsNumeric(v(i, 3)) And Not IsEmpty(v(i, 3)) Then
                    If Not fechas.Exists(d) Then fechas(d) = 1
                    k = cod & "|" & CStr(d)
                    If datos.Exists(k) Then
                        ' dos marcas el mismo dia: se componen
                        datos(k) = (1 + CDbl(datos(k))) * (1 + CDbl(v(i, 3))) - 1
                    Else
                        datos(k) = CDbl(v(i, 3))
                    End If
                    tiene(cod) = 1
                End If
            End If
        End If
    Next i
    If fechas.Count = 0 Then Exit Function

    nFe = fechas.Count
    ReDim fe(1 To nFe)
    j = 0
    For Each kk In fechas.Keys
        j = j + 1
        fe(j) = CDbl(kk)
    Next kk
    OrdenarDoubles fe

    ReDim lista(1 To codigos.Count)
    For Each kk In codigos.Keys
        If tiene.Exists(CStr(kk)) Then
            nL = nL + 1
            lista(nL) = CStr(kk)
        End If
    Next kk
    If nL = 0 Then Exit Function

    ReDim out(1 To nL + 1, 1 To nFe + 2)
    out(1, 1) = "Codigo SBS": out(1, 2) = "Nombre"
    For j = 1 To nFe
        out(1, 2 + j) = fe(j)
    Next j
    For r = 1 To nL
        cod = lista(r)
        out(r + 1, 1) = "'" & cod
        out(r + 1, 2) = nombres(cod)
        For j = 1 To nFe
            k = cod & "|" & CStr(fe(j))
            If datos.Exists(k) Then out(r + 1, 2 + j) = datos(k)
        Next j
    Next r

    Set wsM = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsM.Name = G_MFO
    wsM.Range(wsM.Cells(1, 1), wsM.Cells(nL + 1, nFe + 2)).Value = out
    With wsM
        .Range(.Cells(1, 1), .Cells(1, nFe + 2)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, nFe + 2)).Interior.Color = RGB(64, 64, 64)
        .Range(.Cells(1, 1), .Cells(1, nFe + 2)).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(1, 3), .Cells(1, nFe + 2)).NumberFormat = "dd/mm/yy"
        .Range(.Cells(1, 3), .Cells(1, nFe + 2)).HorizontalAlignment = xlCenter
        .Range(.Cells(2, 3), .Cells(nL + 1, nFe + 2)).NumberFormat = "0.00%"
        .Columns(1).ColumnWidth = 16
        .Columns(2).ColumnWidth = 34
        .Range(.Columns(3), .Columns(nFe + 2)).ColumnWidth = 9
        .Cells(nL + 3, 1).Value = "Var ajustada de cada marca (Marcas_Detalle). Retorno de una ventana = producto de (1 + var) de las marcas con fecha dentro de la ventana, menos 1."
        .Cells(nL + 3, 1).Font.Color = RGB(128, 128, 128)
    End With
    On Error Resume Next
    wsM.Activate
    ActiveWindow.FreezePanes = False
    wsM.Range("C2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    MatrizFondos = nL
End Function


Private Sub OrdenarDoubles(ByRef a() As Double)
    Dim i As Long, j As Long, t As Double
    For i = LBound(a) + 1 To UBound(a)
        t = a(i)
        j = i - 1
        Do While j >= LBound(a)
            If a(j) <= t Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
End Sub


'------------------------------------------------------------------- auxiliares

Private Sub BorrarNombres(wb As Workbook)
    ' Los nombres definidos viajan con la hoja copiada apuntando al libro
    ' original. En la foto no sirven de nada: ya quedaron sustituidos.
    Dim i As Long
    On Error Resume Next
    For i = wb.Names.Count To 1 Step -1
        wb.Names(i).Delete
    Next i
    On Error GoTo 0
End Sub


Private Function CorteDeConfig() As Date
    Dim v As Variant, ws As Worksheet, ult As Long
    CorteDeConfig = Date
    On Error Resume Next
    If HojaExiste(G_CFG) Then
        v = ThisWorkbook.Worksheets(G_CFG).Range("C4").Value
        If IsDate(v) Then
            CorteDeConfig = CDate(v)
            Exit Function
        End If
    End If
    If HojaExiste(G_PRE) Then
        Set ws = ThisWorkbook.Worksheets(G_PRE)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ult >= 2 Then CorteDeConfig = CDate(Application.WorksheetFunction.Max(ws.Range("A2:A" & ult)))
    End If
    On Error GoTo 0
End Function


Private Function InicioDeConfig(dCorte As Date) As Date
    ' Base del FY (Config C7). Vacio: 31/10 del ano anterior al corte.
    Dim v As Variant
    InicioDeConfig = DateSerial(Year(dCorte) - 1, 10, 31)
    On Error Resume Next
    If HojaExiste(G_CFG) Then
        v = ThisWorkbook.Worksheets(G_CFG).Range("C7").Value
        If IsDate(v) Then InicioDeConfig = CDate(v)
    End If
    On Error GoTo 0
End Function


Private Function HojaExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    HojaExiste = Not ws Is Nothing
End Function


Private Function Texto(v As Variant) As String
    If IsError(v) Then
        Texto = ""
    ElseIf IsNull(v) Then
        Texto = ""
    Else
        Texto = Trim$(CStr(v))
    End If
End Function


Private Function SinGuionesG(s As String) As String
    SinGuionesG = Replace(Replace(Trim$(s), "-", ""), " ", "")
End Function


Private Function SinBarraFinal(s As String) As String
    SinBarraFinal = s
    Do While Len(SinBarraFinal) > 0
        If Right$(SinBarraFinal, 1) = "\" Or Right$(SinBarraFinal, 1) = "/" Then
            SinBarraFinal = Left$(SinBarraFinal, Len(SinBarraFinal) - 1)
        Else
            Exit Do
        End If
    Loop
End Function
