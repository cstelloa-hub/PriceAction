Attribute VB_Name = "mod_Guardar"
Option Explicit

'==============================================================================
' mod_Guardar  v3.2  (para mod_Iliquidos v2.70+, un solo libro)
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
'                    F1, F2 y USD mm quedan con formula (SUMPRODUCT) apuntando al
'                    libro del reporte de posiciones (el de Config C75), con la
'                    fecha que uso Cargar posiciones. Funciona con ese libro cerrado.
'     CUADRO_Meses   el mismo cuadro mas un bloque a la derecha con una columna
'                    por mes del ano del corte y el retorno del mes de cada bono
'                    (precio limpio + cupon, como el MTD) y de cada fondo (marcas).
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
Private Const G_PES As String = "BD_Pesos"

' Origen de las posiciones (lo escribe Cargar posiciones en BD_Pesos L6:M18)
Private mExt As String        ' prefijo externo: 'C:\carpeta\[Libro.xlsx]Hoja'!
Private mExtFecha As String   ' DATE(aaaa,mm,dd) de la fecha usada
Private mExtA As Long         ' primera fila de datos en la base
Private mExtB As Long         ' ultima fila
Private mColFe As String, mColCod As String, mColP1 As String, mColP2 As String
Private mColW1 As String, mColW2 As String
Private mFactor As Double
Private mColPeso1 As Long, mColUsd1 As Long   ' columnas J y L del cuadro (se detectan en la fila 4)
Private nExtF As Long
Private Const G_CU2 As String = "CUADRO_Meses"

Private mPaso As String
Private mDiagFo As String


Public Sub GuardarCuadro()
    Dim wbNew As Workbook, wsOri As Worksheet, wsNew As Worksheet, wsNew2 As Worksheet
    Dim carpeta As String, base As String, ruta As String
    Dim dCorte As Date, dIni As Date, n As Long
    Dim scrPrev As Boolean, evPrev As Boolean, alPrev As Boolean, calcPrev As XlCalculation
    Dim t0 As Single, msg As String, nF As Long, nV As Long, nRF As Long, nFo As Long, nMes As Long

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

    ' segundo cuadro, identico, que ademas lleva el bloque mensual de fondos
    mPaso = "copiando el segundo CUADRO"
    wsOri.Copy After:=wsNew
    Set wsNew2 = wbNew.Worksheets(2)
    wsNew2.Name = G_CU2

    ' origen de las posiciones: F1 / F2 / USD mm apuntan al libro del reporte
    mPaso = "leyendo el origen de las posiciones (BD_Pesos L6:M18)"
    LeerOrigenPosiciones wsOri

    ' ---- 2. formulas: las que solo miran la misma hoja se quedan; F1/F2/USD mm
    '         se reescriben contra el libro de posiciones; las que miran
    '         Retornos / Fondos o un nombre definido pasan a valor.
    mPaso = "pasando las formulas del CUADRO"
    CongelarFormulas wsOri, wsNew, nF, nV
    mPaso = "pasando las formulas de " & G_CU2
    CongelarFormulas wsOri, wsNew2, nF, nV
    mPaso = "borrando nombres definidos de la copia"
    BorrarNombres wbNew
    mPaso = "armando el bloque mensual"
    nMes = BloqueMensual(wsOri, wsNew2, dCorte)

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
    Application.Calculate

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
    msg = msg & "Hojas: " & G_CUA & " (sin bloque mensual), " & G_CU2 & " (con bloque mensual), " & G_MRF & ", " & G_MFO & vbCrLf
    msg = msg & "Cuadros: " & nF & " formulas vivas (filas grises), " & nExtF & " celdas F1/F2/USD mm apuntando a " & vbCrLf & _
          "   " & mExt & vbCrLf & nV & " celdas pasadas a valor." & vbCrLf
    msg = msg & "Bloque mensual (" & G_CU2 & "): " & nMes & " retornos de mes (bonos y fondos)." & vbCrLf
    msg = msg & G_MRF & ": " & nRF & " codigos.   " & G_MFO & ": " & nFo & " fondos." & vbCrLf
    If nFo = 0 Then msg = msg & "   NO SALIO " & G_MFO & ": " & mDiagFo & vbCrLf
    If nRF = 0 Then msg = msg & "   (no salio " & G_MRF & ": ningun codigo del cuadro tiene precio en BD_Precios entre esas fechas)" & vbCrLf
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
        ' F1 / F2 / USD mm: contra el libro de posiciones
        f2 = FormulaPeso(wsOri, c)
        If Len(f2) > 0 Then
            wsNew.Range(c.Address).Formula = f2
            nExtF = nExtF + 1
            GoTo Siguiente
        End If
        f2 = SustituirNombres(f, escalares)
        f2 = SustituirCeldasExternas(f2)
        If InStr(f2, "!") > 0 Or TieneNombreRango(f2, rangos) Then
            wsNew.Range(c.Address).Value = c.Value
            nV = nV + 1
        Else
            wsNew.Range(c.Address).Formula = f2
            nF = nF + 1
        End If
Siguiente:
    Next c
End Sub


Private Sub LeerOrigenPosiciones(wsCua As Worksheet)
    ' BD_Pesos L6:M18, escrito por Cargar posiciones (mod_Iliquidos v2.73+).
    Dim ws As Worksheet, ruta As String, hoja As String, carpeta As String, libro As String, d As Date
    Dim j As Long, n As Long
    mExt = "": nExtF = 0
    If Not HojaExiste(G_PES) Then Err.Raise vbObjectError + 110, , "No existe BD_Pesos. Corre Cargar posiciones."
    Set ws = ThisWorkbook.Worksheets(G_PES)
    If Texto(ws.Range("L6").Value) <> "ORIGEN (para Guardar la foto)" Then Err.Raise vbObjectError + 111, , _
        "BD_Pesos no tiene el bloque ORIGEN (L6:M18). Corre 2. Cargar posiciones con mod_Iliquidos v2.73 o mas nuevo."
    ruta = Texto(ws.Range("M7").Value)
    hoja = Texto(ws.Range("M8").Value)
    mExtA = CLng(ws.Range("M9").Value) + 1
    mExtB = CLng(ws.Range("M10").Value)
    d = CDate(ws.Range("M11").Value)
    mColFe = Texto(ws.Range("M12").Value)
    mColCod = Texto(ws.Range("M13").Value)
    mColP1 = Texto(ws.Range("M14").Value)
    mColP2 = Texto(ws.Range("M15").Value)
    mColW1 = Texto(ws.Range("M16").Value)
    mColW2 = Texto(ws.Range("M17").Value)
    mFactor = CDbl(ws.Range("M18").Value)
    If mFactor <= 0 Then mFactor = 1
    If Len(ruta) = 0 Or Len(hoja) = 0 Or mExtB < mExtA Then Err.Raise vbObjectError + 112, , _
        "El bloque ORIGEN de BD_Pesos esta incompleto. Vuelve a correr 2. Cargar posiciones."
    carpeta = Left$(ruta, InStrRev(ruta, Application.PathSeparator))
    libro = Mid$(ruta, InStrRev(ruta, Application.PathSeparator) + 1)
    mExt = "'" & carpeta & "[" & libro & "]" & hoja & "'!"
    mExtFecha = "DATE(" & Year(d) & "," & Month(d) & "," & Day(d) & ")"

    ' columnas del cuadro: en la fila 4 hay dos pares "F1","F2": pesos y USD mm
    mColPeso1 = 0: mColUsd1 = 0
    For j = 1 To 40
        If UCase$(Texto(wsCua.Cells(4, j).Value)) = "F1" Then
            n = n + 1
            If n = 1 Then mColPeso1 = j
            If n = 2 Then mColUsd1 = j
        End If
    Next j
    If mColPeso1 = 0 Or mColUsd1 = 0 Then Err.Raise vbObjectError + 113, , _
        "No encontre las columnas F1 (peso y USD mm) en la fila 4 del CUADRO."
End Sub


Private Function FormulaPeso(wsCua As Worksheet, c As Range) As String
    ' Devuelve la formula externa para una celda F1 / F2 / USD mm; "" si la
    ' celda no es de esas (o es una fila gris, que se maneja por su propia suma).
    Dim r As Long, col As Long, cod As String, cond As String, rangoDato As String
    Dim fPeso As String, re As Object
    r = c.Row: col = c.Column
    If r < 5 Then Exit Function
    If col <> mColPeso1 And col <> mColPeso1 + 1 And col <> mColUsd1 And col <> mColUsd1 + 1 Then Exit Function
    cod = Texto(wsCua.Cells(r, 2).Value)
    ' fila manual (peso tipeado, sin formula en F1): se deja como esta
    If Len(cod) > 0 And Not wsCua.Cells(r, mColPeso1).HasFormula Then Exit Function

    If Len(cod) > 0 Then
        ' fila de instrumento: SUMPRODUCT contra el libro de posiciones.
        ' SUMPRODUCT (y no SUMIFS) porque funciona con el libro cerrado.
        cond = "(SUBSTITUTE(" & Ext(mColCod) & ",""-"","""")=$B" & r & ")*(" & Ext(mColFe) & "=" & mExtFecha & ")"
        If col = mColPeso1 Then
            FormulaPeso = "=IFERROR(SUMPRODUCT(" & cond & "," & Ext(mColW1) & ")/" & NumTexto(mFactor) & ",0)"
        ElseIf col = mColPeso1 + 1 Then
            FormulaPeso = "=IFERROR(SUMPRODUCT(" & cond & "," & Ext(mColW2) & ")/" & NumTexto(mFactor) & ",0)"
        ElseIf col = mColUsd1 Then
            FormulaPeso = "=IFERROR(SUMPRODUCT(" & cond & "," & Ext(mColP1) & ")/$" & ColLetra(mColUsd1) & "$1/1000000,"""")"
        ElseIf col = mColUsd1 + 1 Then
            FormulaPeso = "=IFERROR(SUMPRODUCT(" & cond & "," & Ext(mColP2) & ")/$" & ColLetra(mColUsd1) & "$1/1000000,"""")"
        End If
    Else
        ' fila gris: USD mm = la misma suma que su F1 / F2 pero sobre las
        ' columnas de USD mm (J->L, K->M). F1 / F2 ya son de la misma hoja.
        If col = mColUsd1 Or col = mColUsd1 + 1 Then
            fPeso = wsCua.Cells(r, col - (mColUsd1 - mColPeso1)).Formula
            If Len(fPeso) > 1 Then
                FormulaPeso = MoverColumna(fPeso, ColLetra(col - (mColUsd1 - mColPeso1)), ColLetra(col))
            End If
        End If
    End If
End Function


Private Function Ext(letra As String) As String
    Ext = mExt & "$" & letra & "$" & mExtA & ":$" & letra & "$" & mExtB
End Function


Private Function MoverColumna(f As String, de As String, a As String) As String
    ' Cambia las referencias de la columna "de" a la columna "a" en una formula
    ' de la misma hoja: $J$7 -> $L$7, J7 -> L7. No toca nombres de funciones.
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(^|[^A-Za-z0-9_$])(\$?)" & de & "(\$?\d+)"
    MoverColumna = re.Replace(f, "$1$2" & a & "$3")
End Function


Private Function ColLetra(n As Long) As String
    Dim s As String, k As Long
    k = n
    Do While k > 0
        s = Chr$(65 + (k - 1) Mod 26) & s
        k = (k - 1) \ 26
    Loop
    ColLetra = s
End Function


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


'------------------------------------------------------------ bloque mensual

Private Function BloqueMensual(wsOri As Worksheet, wsNew As Worksheet, dCorte As Date) As Long
    ' A la derecha del cuadro (dos columnas despues de la ultima), una columna
    ' por mes del ano del corte con el RETORNO DEL MES de cada fila:
    '   bonos:  (P fin de mes / P fin del mes anterior - 1) + cupon x dias / 365 / P base
    '           (+ pase USD->PEN de Config C58 prorrateado, si es USD), igual que la
    '           columna MTD del cuadro. Mes en curso = hasta el corte.
    '   fondos: producto de (1 + var ajustada) de las marcas del mes, menos 1.
    ' Vacio: sin precio al inicio o al fin del mes, sin marca, fila gris, mes futuro.
    Dim ws As Worksheet, v As Variant, ult As Long, i As Long, cod As String, d As Double
    Dim acum As Object, precios As Object, cupones As Object, k As String
    Dim c0 As Long, m As Long, r As Long, ultFila As Long, ultCol As Long
    Dim anio As Long, n As Long, rg As Range, spread As Double, mon As String
    Dim dIniMes As Double, dFinMes As Double, dF As Double, dB As Double, pF As Double, pB As Double, cup As Double
    Dim x As Double, hay As Boolean

    anio = Year(dCorte)
    On Error Resume Next
    spread = CDbl(ThisWorkbook.Worksheets(G_CFG).Range("C58").Value)
    On Error GoTo 0

    ' --- precios: codigo|fecha -> precio limpio (col E) y cupon (col G), todo el historico
    Set precios = CreateObject("Scripting.Dictionary"): precios.CompareMode = 1
    Set cupones = CreateObject("Scripting.Dictionary"): cupones.CompareMode = 1
    If HojaExiste(G_PRE) Then
        Set ws = ThisWorkbook.Worksheets(G_PRE)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ult >= 2 Then
            v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 7)).Value2
            For i = 1 To UBound(v, 1)
                If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
                    cod = SinGuionesG(Texto(v(i, 2)))
                    If Len(cod) > 0 And IsNumeric(v(i, 5)) And Not IsEmpty(v(i, 5)) Then
                        k = cod & "|" & CStr(CDbl(v(i, 1)))
                        precios(k) = CDbl(v(i, 5))
                        If IsNumeric(v(i, 7)) And Not IsEmpty(v(i, 7)) Then cupones(k) = CDbl(v(i, 7))
                    End If
                End If
            Next i
        End If
    End If

    ' --- marcas: codigo|mes -> producto acumulado de (1 + var)
    Set acum = CreateObject("Scripting.Dictionary"): acum.CompareMode = 1
    If HojaExiste(G_MDE) Then
        Set ws = ThisWorkbook.Worksheets(G_MDE)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ult >= 2 Then
            v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 4)).Value2
            For i = 1 To UBound(v, 1)
                cod = SinGuionesG(Texto(v(i, 1)))
                If Len(cod) > 0 And IsNumeric(v(i, 3)) And Not IsEmpty(v(i, 3)) Then
                    d = CDbl(v(i, 3))
                    If Year(d) = anio And d <= CDbl(dCorte) Then
                        If IsNumeric(v(i, 4)) And Not IsEmpty(v(i, 4)) Then
                            k = cod & "|" & Month(d)
                            If acum.Exists(k) Then
                                acum(k) = CDbl(acum(k)) * (1 + CDbl(v(i, 4)))
                            Else
                                acum(k) = 1 + CDbl(v(i, 4))
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    End If

    ultFila = wsOri.Cells(wsOri.Rows.Count, 1).End(xlUp).Row
    If ultFila < 5 Then Exit Function
    ultCol = wsOri.Cells(4, wsOri.Columns.Count).End(xlToLeft).Column    ' fila 4 = cabecera del cuadro
    c0 = ultCol + 2

    With wsNew
        .Cells(3, c0).Value = "Retorno del mes. Bonos: precio limpio fin de mes vs fin del mes anterior + cupon devengado" & _
                              IIf(spread <> 0, " (+ pase USD->PEN)", "") & ", como el MTD. Fondos: marcas del mes encadenadas. Vacio = sin dato."
        .Cells(3, c0).Font.Color = RGB(128, 128, 128)
        .Cells(3, c0).Font.Size = 7
        For m = 1 To 12
            .Cells(4, c0 + m - 1).Value = Format$(DateSerial(anio, m, 1), "mm/yy")
        Next m
        With .Range(.Cells(4, c0), .Cells(4, c0 + 11))
            .Interior.Color = RGB(192, 0, 0)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        .Columns(c0 - 1).ColumnWidth = 1.6
        .Range(.Columns(c0), .Columns(c0 + 11)).ColumnWidth = 8.5

        For r = 5 To ultFila
            cod = SinGuionesG(Texto(wsOri.Cells(r, 2).Value))
            mon = UCase$(Texto(wsOri.Cells(r, 3).Value))
            If Len(cod) = 0 Then
                ' fila gris / etiqueta: se copia el color para que la banda siga
                If wsOri.Cells(r, 1).Interior.ColorIndex <> xlNone Then
                    .Range(.Cells(r, c0), .Cells(r, c0 + 11)).Interior.Color = wsOri.Cells(r, 1).Interior.Color
                End If
            Else
                For m = 1 To 12
                    dIniMes = CDbl(DateSerial(anio, m, 1))
                    If dIniMes > CDbl(dCorte) Then Exit For
                    hay = False
                    k = cod & "|" & m
                    If acum.Exists(k) Then
                        ' fondo
                        x = CDbl(acum(k)) - 1
                        hay = True
                    Else
                        ' bono: ultimo precio <= fin de mes (dentro del mes) y ultimo <= dia antes del inicio
                        dFinMes = CDbl(DateSerial(anio, m + 1, 0))
                        If dFinMes > CDbl(dCorte) Then dFinMes = CDbl(dCorte)
                        dF = UltimoPrecio(precios, cod, dFinMes, dIniMes)
                        dB = UltimoPrecio(precios, cod, dIniMes - 1, dIniMes - 15)
                        If dF > 0 And dB > 0 Then
                            pF = CDbl(precios(cod & "|" & CStr(dF)))
                            pB = CDbl(precios(cod & "|" & CStr(dB)))
                            cup = 0
                            If cupones.Exists(cod & "|" & CStr(dF)) Then cup = CDbl(cupones(cod & "|" & CStr(dF)))
                            If pB > 0 Then
                                x = (pF - pB) / pB + cup * (dF - dB) / 365 / pB
                                If mon = "USD" Then x = x + spread * (dF - dB) / 365
                                hay = True
                            End If
                        End If
                    End If
                    If hay Then
                        .Cells(r, c0 + m - 1).Value = x
                        n = n + 1
                    End If
                Next m
            End If
        Next r

        Set rg = .Range(.Cells(5, c0), .Cells(ultFila, c0 + 11))
        rg.NumberFormat = "0.00%"
        rg.Font.Size = 8
        rg.HorizontalAlignment = xlCenter
        With rg.Borders
            .LineStyle = xlContinuous
            .Color = RGB(217, 217, 217)
            .Weight = xlThin
        End With
    End With
    BloqueMensual = n
End Function


Private Function UltimoPrecio(precios As Object, cod As String, dHasta As Double, dDesde As Double) As Double
    ' Fecha del ultimo precio del codigo entre dDesde y dHasta (0 si no hay).
    Dim dd As Double
    For dd = dHasta To dDesde Step -1
        If precios.Exists(cod & "|" & CStr(dd)) Then
            UltimoPrecio = dd
            Exit Function
        End If
    Next dd
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
    ' Marcas_Detalle: A codigo  B nombre en Marcas  C fecha  D var ajustada  E ln
    Dim ws As Worksheet, v As Variant, ult As Long, i As Long
    Dim codigos As Object, nombres As Object, fechas As Object, datos As Object, tiene As Object
    Dim cod As String, d As Double, k As String
    Dim fe() As Double, nFe As Long, j As Long, out() As Variant, r As Long
    Dim wsM As Worksheet, kk As Variant, lista() As String, nL As Long

    Dim nFil As Long, nFecha As Long, nRango As Long, nCod As Long, ejMd As String, ejIn As String
    mDiagFo = ""
    If Not HojaExiste(G_MDE) Then
        mDiagFo = "no existe la hoja " & G_MDE
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(G_MDE)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then
        mDiagFo = G_MDE & " esta vacia: corre Cargar Marcas"
        Exit Function
    End If
    v = ws.Range(ws.Cells(2, 1), ws.Cells(ult, 4)).Value2

    Set codigos = CodigosDelCuadro(nombres)
    Set fechas = CreateObject("Scripting.Dictionary")
    Set datos = CreateObject("Scripting.Dictionary"): datos.CompareMode = 1
    Set tiene = CreateObject("Scripting.Dictionary"): tiene.CompareMode = 1

    For i = 1 To UBound(v, 1)
        cod = SinGuionesG(Texto(v(i, 1)))
        If Len(cod) > 0 Then
            nFil = nFil + 1
            If nFil <= 3 Then ejMd = ejMd & "[" & cod & " | " & Texto(v(i, 3)) & " | " & Texto(v(i, 4)) & "] "
        End If
        If Len(cod) > 0 And IsNumeric(v(i, 3)) And Not IsEmpty(v(i, 3)) Then
            nFecha = nFecha + 1
            d = CDbl(v(i, 3))
            If d >= CDbl(dIni) And d <= CDbl(dCorte) Then nRango = nRango + 1
            If codigos.Exists(cod) Then nCod = nCod + 1
            If d >= CDbl(dIni) And d <= CDbl(dCorte) And codigos.Exists(cod) Then
                If IsNumeric(v(i, 4)) And Not IsEmpty(v(i, 4)) Then
                    If Not fechas.Exists(d) Then fechas(d) = 1
                    k = cod & "|" & CStr(d)
                    If datos.Exists(k) Then
                        ' dos marcas el mismo dia: se componen
                        datos(k) = (1 + CDbl(datos(k))) * (1 + CDbl(v(i, 4))) - 1
                    Else
                        datos(k) = CDbl(v(i, 4))
                    End If
                    tiene(cod) = 1
                End If
            End If
        End If
    Next i
    If fechas.Count = 0 Then
        j = 0
        For Each kk In codigos.Keys
            j = j + 1
            If j <= 3 Then ejIn = ejIn & "[" & CStr(kk) & "] "
        Next kk
        mDiagFo = "filas en " & G_MDE & ": " & nFil & " | con fecha numerica en col C: " & nFecha & _
                  " | dentro de " & Format$(dIni, "dd/mm/yy") & " - " & Format$(dCorte, "dd/mm/yy") & ": " & nRango & _
                  " | cuyo codigo esta en el cuadro: " & nCod & vbCrLf & _
                  "   primeras filas de " & G_MDE & " (codigo | col C | col D): " & ejMd & vbCrLf & _
                  "   primeros codigos del cuadro: " & ejIn
        Exit Function
    End If

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

Private Sub AValores(ws As Worksheet)
    Dim rg As Range
    On Error Resume Next
    Set rg = ws.UsedRange
    On Error GoTo 0
    If rg Is Nothing Then Exit Sub
    rg.Value = rg.Value
End Sub


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
