Attribute VB_Name = "mod_PriceAction"
Option Explicit

'==============================================================================
' mod_PriceAction   -   PriceAction.xlsm
'
' EL PRODUCTO: "Retornos: Cuasi Sob & Iliquidos". Junta tres fuentes:
'
'   Base_RF.xlsm       precios del vector SBS + maestro de bonos   (solo lectura)
'   Base_Marcas.xlsm   marcas de fondos + maestro + datos          (solo lectura)
'   FMS                posicionamiento F1 / F2                     (solo lectura)
'
' Este libro NO escribe nunca en las bases. Las abre, se copia lo que
' necesita a hojas propias (Precios, Instrumentos, Marcas, Datos_Fondos) y
' las cierra. Esas copias son la materia prima del cuadro y quedan en el
' libro para que cada numero se pueda rastrear hasta su origen.
'
' Hojas:
'   Config         rutas, corte, ventanas, spread, prima, niveles, mapa FMS
'   BD_Pesos       <- FMS. Peso de cada codigo en F1 y F2
'   Manual         filas a mano (Deals exPeru) que sobreviven a cada corrida
'   Manual_Var     marcas de esas filas a mano
'   Precios        copia de BD_Precios de Base_RF
'   Instrumentos   los dos maestros juntos, con el nombre final armado
'   Marcas         copia de Marcas_Detalle de Base_Marcas, con las ventanas
'   Datos_Fondos   copia del historico de duracion / spread / CY
'   Retornos       calculo por bono y ventana, con precios y dias a la vista
'   Fondos         calculo por fondo y ventana, con chequeo y conteo
'   CUADRO         el producto final
'   Panel          los botones
'   Log            el ultimo error, completo
'
' Rutina del dia:   1. Cargar FMS    2. ARMAR CUADRO    3. Guardar foto
'==============================================================================

Private Const VER As String = "3.5"

Private Const SH_CFG As String = "Config"
Private Const SH_PES As String = "BD_Pesos"
Private Const SH_MAN As String = "Manual"
Private Const SH_MVA As String = "Manual_Var"
Private Const SH_PRE As String = "Precios"
Private Const SH_INS As String = "Instrumentos"
Private Const SH_MDE As String = "Marcas"
Private Const SH_DAT As String = "Datos_Fondos"
Private Const SH_RET As String = "Retornos"
Private Const SH_FON As String = "Fondos"
Private Const SH_CUA As String = "CUADRO"
Private Const SH_LOG As String = "Log"
Private Const SH_PAN As String = "Panel"

' BD_Pesos
Private Const PE_COLS As Long = 10
' 1 Codigo SBS  2 Nemonico  3 Emisor  4 Moneda  5 Val_total F1  6 Val_total F2
' 7 Peso F1  8 Peso F2  9 Fecha FMS  10 Vencimiento (FMS)
' M1 M2 M3: total F1, total F2 (soles) y tipo de cambio del FMS

' Precios (copia de Base_RF.BD_Precios)
Private Const PR_COLS As Long = 10

' Instrumentos (los dos maestros juntos)
'   A Codigo  B Nombre final  C Categoria  D Fuera/REF  E Moneda  F CY (fondos)
'   G ISIN  H Rating  I Duracion  J Spread  K Origen  L Nombre base  M Cupon
'   N Vencimiento  O Fecha datos (fondos)
Private Const IN_COLS As Long = 15

Private mT0 As Double
Private mPasos As String
Private mDetalle As String
Private mSilencio As Boolean
Private mResumen As String
Private mFallo As Boolean
Private mForzar As Boolean

' CUADRO: filas fijas
Private Const CU_FIN   As Long = 2
Private Const CU_BASE  As Long = 3
Private Const CU_CAB   As Long = 4
Private Const CU_DAT   As Long = 5

' CUADRO: columnas (todas visibles, ninguna de ayuda)
Private Const CC_NOM   As Long = 1
Private Const CC_COD   As Long = 2
Private Const CC_MRE   As Long = 3
Private Const CC_ISIN  As Long = 4
Private Const CC_RAT   As Long = 5
Private Const CC_DUR   As Long = 6
Private Const CC_SPR   As Long = 7
Private Const CC_CY    As Long = 8
Private Const CC_CYP   As Long = 9
Private Const CC_P1    As Long = 10
Private Const CC_U1    As Long = 12
Private Const CC_SEP   As Long = 14
Private Const CC_W1    As Long = 15
Private Const CC_SEP2  As Long = 21
Private Const CC_MEN   As Long = 22
Private Const CC_ANN   As Long = 23
Private Const N_VENT   As Long = 6
Private Const CAT_FON  As String = "Fondos Tradicionales"


'==============================================================================
'  CONFIG
'==============================================================================
Public Sub CrearConfig()
    Dim ws As Worksheet, prev As Object, k As Variant, i As Long
    Dim scrPrev As Boolean

    scrPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Falla

    Set prev = NuevoDic()
    On Error Resume Next
    For i = 2 To 42
        prev("C" & i) = ThisWorkbook.Worksheets(SH_CFG).Range("C" & i).Value
    Next i
    On Error GoTo Falla

    Set ws = HojaOCrea(SH_CFG)
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    ws.Range("B1").Value = "PRICE ACTION  -  configuracion"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 13

    Cfg ws, 2, "Base_RF.xlsm (ruta completa)", Rec(prev, "C2", ""), "Se abre en solo lectura."
    Cfg ws, 3, "Base_Marcas.xlsm (ruta completa)", Rec(prev, "C3", ""), "Se abre en solo lectura."
    Cfg ws, 4, "Carpeta del FMS", Rec(prev, "C4", ""), "Donde estan los 'FMS_AAAAMMDD.xlsx'."
    Cfg ws, 5, "Fecha de corte", Rec(prev, "C5", Empty), _
        "Vacio = el ultimo precio que haya en Base_RF."
    Cfg ws, 6, "Inicio del periodo libre (MayoTD)", Rec(prev, "C6", DateSerial(2026, 5, 1)), _
        "La base del periodo es esta fecha menos un dia."
    Cfg ws, 7, "Base del FY", Rec(prev, "C7", DateSerial(2025, 10, 31)), "FY = 01 nov a 31 oct."
    Cfg ws, 8, "Dias de la ventana 5D", Rec(prev, "C8", 5), "Dias HABILES hacia atras."
    Cfg ws, 9, "Dias sin precio antes de REVISAR", Rec(prev, "C9", 5), _
        "Si la base o el fin usan un vector mas viejo que esto, lo marca."
    Cfg ws, 10, "Codigo del Fondo 1", Rec(prev, "C10", "01"), "Columna A de la hoja Cartera del FMS."
    Cfg ws, 11, "Codigo del Fondo 2", Rec(prev, "C11", "02"), ""
    Cfg ws, 12, "Total Fondo 1 a mano", Rec(prev, "C12", Empty), "Solo si el Anexo I falla. Vacio = se lee del Anexo I."
    Cfg ws, 13, "Total Fondo 2 a mano", Rec(prev, "C13", Empty), "En soles."
    Cfg ws, 14, "Usar marcas posteriores al corte", Rec(prev, "C14", "SI"), _
        "SI = un fondo que marco despues del corte entra con esa marca. NO = corta en el corte."
    Cfg ws, 15, "Spread USD -> PEN (anual)", Rec(prev, "C15", 0.0133), _
        "Lo que esta en USD pasa a soles sumando esto, prorrateado por dias. 0 = no convierte."
    Map ws, 16, "Categorias que van siempre en PEN", Txt(Rec(prev, "C16", "Cuasi Soberanos"))
    Map ws, 17, "Orden de las categorias", Txt(Rec(prev, "C17", "Cuasi Soberanos|Iliquidos|Coinversiones|Fondos Tradicionales"))
    Map ws, 18, "SubTotal Fondos: categorias", Txt(Rec(prev, "C18", "Fondos Tradicionales"))
    Map ws, 19, "SubTotal Fondos + Coinversiones: categorias", Txt(Rec(prev, "C19", "Fondos Tradicionales|Coinversiones"))
    ws.Range("D16").Value = "Aunque el instrumento este en USD. Separa con |"
    ws.Range("D17").Value = "De mas liquido a mas iliquido; los fondos al final. Separa con |. Las que no esten aqui van al final."
    ws.Range("D18").Value = "Separa con |. Vacio = la fila no sale."
    ws.Range("D19").Value = "Separa con |. Vacio = la fila no sale."
    ws.Range("D16:D19").Font.Color = RGB(128, 128, 128)
    Cfg ws, 20, "Nivel de cartera objetivo PEN", Rec(prev, "C20", 0.1), "Solo sale en el titulo de la fila."
    Cfg ws, 21, "Nivel de cartera objetivo USD", Rec(prev, "C21", 0.09), "Idem."
    Cfg ws, 22, "Prima en SubTotales y Total (pb)", Rec(prev, "C22", 100), _
        "Se SUMA a las columnas de retorno de SubTotal Fondos, SubTotal Fondos + Coinv. y Total. 0 = sin prima."
    Cfg ws, 23, "Prima: PLANO o PRORRATEADO", Txt(Rec(prev, "C23", "PRORRATEADO")), _
        "PRORRATEADO = 100 pb anuales repartidos por los dias de cada ventana. PLANO = +1.00% parejo."
    Cfg ws, 24, "Tipo de cambio USD/PEN a mano", Rec(prev, "C24", Empty), _
        "VACIO = usa el del FMS (hoja T.Cambio). Solo escribe aqui para pisarlo."
    Cfg ws, 25, "Carpeta de la foto (Guardar foto)", Rec(prev, "C25", ""), _
        "Vacio = subcarpeta Resumenes junto a este libro. La usa mod_Guardar."
    ws.Range("C2:C25").Interior.Color = RGB(255, 242, 204)
    ws.Range("C5:C7").NumberFormat = "dd/mm/yyyy"
    ws.Range("C15").NumberFormat = "0.00%"
    ws.Range("C20:C21").NumberFormat = "0.0%"

    ws.Range("B26").Value = "MAPA DEL FMS  (letras de columna y nombres de hoja)"
    ws.Range("B26").Font.Bold = True
    Map ws, 27, "Cartera: Codigo fondo", Txt(Rec(prev, "C27", "A"))
    Map ws, 28, "Cartera: Asset class", Txt(Rec(prev, "C28", "C"))
    Map ws, 29, "Cartera: Codigo SBS", Txt(Rec(prev, "C29", "D"))
    Map ws, 30, "Cartera: Emisor", Txt(Rec(prev, "C30", "E"))
    Map ws, 31, "Cartera: Nemonico", Txt(Rec(prev, "C31", "G"))
    Map ws, 32, "Cartera: Moneda", Txt(Rec(prev, "C32", "H"))
    Map ws, 33, "Cartera: Val_total", Txt(Rec(prev, "C33", "J"))
    Map ws, 34, "Cartera: Fecha vencimiento", Txt(Rec(prev, "C34", "L"))
    Map ws, 35, "Hoja de la cartera", Txt(Rec(prev, "C35", "Cartera"))
    Map ws, 36, "Hoja del anexo", Txt(Rec(prev, "C36", "Anexo I"))
    Map ws, 37, "Codigo del total en el anexo", Txt(Rec(prev, "C37", "1.1.1"))
    Map ws, 38, "Anexo: Codigo fondo", Txt(Rec(prev, "C38", "A"))
    Map ws, 39, "Anexo: Codigo", Txt(Rec(prev, "C39", "D"))
    Map ws, 40, "Anexo: Monto", Txt(Rec(prev, "C40", "F"))
    Map ws, 41, "Hoja del tipo de cambio", Txt(Rec(prev, "C41", "T.Cambio"))
    Map ws, 42, "Simbolo del dolar en esa hoja", Txt(Rec(prev, "C42", "US$"))
    ws.Range("C27:C42").Interior.Color = RGB(255, 242, 204)
    ws.Range("C27:C34").HorizontalAlignment = xlCenter
    ws.Range("C38:C40").HorizontalAlignment = xlCenter

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 48
    ws.Columns("D").ColumnWidth = 90

    Application.ScreenUpdating = scrPrev
    MsgBox "Config lista (v" & VER & ")." & vbCrLf & vbCrLf & _
           "Pon las rutas de las dos bases (C2, C3) y la carpeta del FMS (C4).", _
           vbInformation, "CrearConfig"
    Exit Sub

Falla:
    Application.ScreenUpdating = True
    MsgBox "CrearConfig se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "CrearConfig"
End Sub


Private Function CorteEfectivo(wsCfg As Worksheet) As Double
    ' Config C5. Vacio: la ultima fecha de la hoja Precios (copia de la base);
    ' si tampoco hay, hoy.
    Dim ws As Worksheet, ult As Long
    If IsDate(wsCfg.Range("C5").Value) Then
        CorteEfectivo = CDbl(CDate(wsCfg.Range("C5").Value))
        Exit Function
    End If
    If HojaExiste(SH_PRE) Then
        Set ws = ThisWorkbook.Worksheets(SH_PRE)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ult >= 2 Then
            CorteEfectivo = Application.WorksheetFunction.Max(ws.Range("A2:A" & ult))
            If CorteEfectivo > 0 Then Exit Function
        End If
    End If
    CorteEfectivo = CDbl(Date)
End Function


'==============================================================================
'  CARGAR FMS  ->  BD_Pesos
'==============================================================================
Public Sub CargarFMS()
    Dim wsCfg As Worksheet, wb As Workbook, wsCar As Worksheet, wsAnx As Worksheet
    Dim carpeta As String, ruta As String, hojaCar As String, hojaAnx As String
    Dim f1 As String, f2 As String, codTot As String
    Dim dCorte As Double, dFms As Double, tcFms As Double
    Dim v As Variant, va As Variant
    Dim cMap As Object, aMap As Object, datos As Object
    Dim totF1 As Double, totF2 As Double, aMano As Boolean
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
    carpeta = SinBarra(Txt(wsCfg.Range("C4").Value))
    If Len(carpeta) = 0 Then Err.Raise vbObjectError + 1, , "Falta la carpeta del FMS en Config C4."
    hojaCar = Txt(wsCfg.Range("C35").Value): If Len(hojaCar) = 0 Then hojaCar = "Cartera"
    hojaAnx = Txt(wsCfg.Range("C36").Value): If Len(hojaAnx) = 0 Then hojaAnx = "Anexo I"
    codTot = Txt(wsCfg.Range("C37").Value): If Len(codTot) = 0 Then codTot = "1.1.1"
    f1 = NormFondo(Txt(wsCfg.Range("C10").Value))
    f2 = NormFondo(Txt(wsCfg.Range("C11").Value))

    dCorte = CorteEfectivo(wsCfg)

    Paso "Buscando el ultimo FMS <= " & Format$(CDate(dCorte), "dd/mm/yyyy")
    dFms = BuscarFms(carpeta, dCorte, ruta)
    If dFms = 0 Then Err.Raise vbObjectError + 2, , _
        "No encontre ningun 'FMS_AAAAMMDD.xlsx' en los 20 dias previos a " & _
        Format$(CDate(dCorte), "dd/mm/yyyy") & "." & vbCrLf & vbCrLf & QueHayEnCarpeta(carpeta)

    If Not mForzar Then
        If FmsYaCargado(dFms) Then
            Restaurar calcPrev, scrPrev, evPrev
            Aviso "BD_Pesos ya tiene el FMS del " & Format$(CDate(dFms), "dd/mm/yyyy") & _
                  ", que es el ultimo <= corte." & vbCrLf & "No se volvio a abrir." & vbCrLf & vbCrLf & _
                  "Si quieres releerlo igual: Cargar FMS FORZADO.", "CargarFMS"
            Exit Sub
        End If
    End If

    Paso "Abriendo el FMS " & Format$(CDate(dFms), "dd/mm/yyyy") & " en solo lectura"
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    On Error Resume Next
    Set wsCar = wb.Worksheets(hojaCar)
    Set wsAnx = wb.Worksheets(hojaAnx)
    On Error GoTo Falla
    If wsCar Is Nothing Then
        wb.Close SaveChanges:=False
        Err.Raise vbObjectError + 3, , "No encontre la hoja '" & hojaCar & "' en el FMS."
    End If

    Paso "Leyendo la cartera a memoria"
    Set cMap = MapaFms(wsCfg)
    v = LeerHoja(wsCar, MaxCol(cMap))
    If Not wsAnx Is Nothing Then
        Set aMap = NuevoDic()
        aMap("fondo") = Letra(Txt(wsCfg.Range("C38").Value), 1)
        aMap("codigo") = Letra(Txt(wsCfg.Range("C39").Value), 4)
        aMap("monto") = Letra(Txt(wsCfg.Range("C40").Value), 6)
        va = LeerHoja(wsAnx, MaxCol(aMap))
    End If
    Paso "Leyendo el tipo de cambio de la hoja " & Txt(wsCfg.Range("C41").Value)
    tcFms = TipoCambioFms(wb, wsCfg)
    wb.Close SaveChanges:=False
    Set wb = Nothing: Set wsCar = Nothing: Set wsAnx = Nothing

    Paso "Sumando Val_total por codigo"
    ' TODOS los codigos del FMS, sin filtro. El cuadro luego toma solo los
    ' que esten en los maestros; el resto queda aqui para poder revisar
    ' "este codigo esta en el fondo y no lo estoy siguiendo".
    Set datos = NuevoDic()
    AcumularCartera v, cMap, f1, 1, datos
    AcumularCartera v, cMap, f2, 2, datos
    If datos.Count = 0 Then Err.Raise vbObjectError + 4, , _
        "No encontre filas de los fondos " & f1 & " / " & f2 & " en la hoja " & hojaCar & "." & vbCrLf & _
        "Revisa Config C10, C11 y C27."

    Paso "Leyendo el total de cada fondo del Anexo I"
    totF1 = Num(wsCfg.Range("C12").Value)
    totF2 = Num(wsCfg.Range("C13").Value)
    If totF1 > 0 Or totF2 > 0 Then aMano = True
    If totF1 <= 0 Then totF1 = TotalAnexo(va, aMap, f1, codTot)
    If totF2 <= 0 Then totF2 = TotalAnexo(va, aMap, f2, codTot)
    If totF1 <= 0 Then totF1 = SumaDic(datos, 1)
    If totF2 <= 0 Then totF2 = SumaDic(datos, 2)

    Paso "Escribiendo BD_Pesos"
    EscribirPesos datos, totF1, totF2, dFms
    GuardarTC tcFms

    Restaurar calcPrev, scrPrev, evPrev
    Aviso "v" & VER & " - CargarFMS terminado en " & Format$(Timer - mT0, "0.0") & " s." & vbCrLf & vbCrLf & _
          mPasos & vbCrLf & _
          "FMS usado: " & Format$(CDate(dFms), "dd/mm/yyyy") & vbCrLf & _
          "Codigos en BD_Pesos: " & datos.Count & vbCrLf & _
          "Total Fondo 1: " & Format$(totF1 / 1000000, "#,##0.0") & " mn" & _
          IIf(aMano, "  (a mano, Config C12)", "") & vbCrLf & _
          "Total Fondo 2: " & Format$(totF2 / 1000000, "#,##0.0") & " mn" & _
          IIf(aMano, "  (a mano, Config C13)", "") & vbCrLf & _
          "Tipo de cambio: " & IIf(tcFms > 0, Format$(tcFms, "0.0000"), _
          "NO LO ENCONTRE - revisa Config C41 y C42") & vbCrLf & vbCrLf & _
          "Ahora: ARMAR CUADRO.", "CargarFMS"
    Exit Sub

Falla:
    Reventar wb, calcPrev, scrPrev, evPrev, "CargarFMS"
End Sub


Private Function QueHayEnCarpeta(carpeta As String) As String
    ' Para el mensaje de error: si la carpeta existe y que archivos FMS* ve.
    Dim nom As String, n As Long, lista As String
    QueHayEnCarpeta = "Carpeta (Config C4): " & carpeta & vbCrLf
    If Len(Dir(carpeta, vbDirectory)) = 0 Then
        QueHayEnCarpeta = QueHayEnCarpeta & "ESA CARPETA NO EXISTE o no tienes acceso. Revisa la ruta."
        Exit Function
    End If
    On Error Resume Next
    nom = Dir(carpeta & Application.PathSeparator & "*FMS*.*")
    Do While Len(nom) > 0 And n < 8
        n = n + 1
        lista = lista & "   " & nom & vbCrLf
        nom = Dir
    Loop
    On Error GoTo 0
    If n = 0 Then
        QueHayEnCarpeta = QueHayEnCarpeta & "La carpeta existe pero no tiene ningun archivo con 'FMS' en el nombre."
    Else
        QueHayEnCarpeta = QueHayEnCarpeta & "Archivos con 'FMS' que si veo ahi (hasta 8):" & vbCrLf & lista & _
            "El nombre tiene que ser exactamente FMS_AAAAMMDD.xlsx (o .xls / .xlsb)."
    End If
End Function


Private Sub AcumularCartera(v As Variant, cMap As Object, fondo As String, _
                            cual As Long, datos As Object)
    Dim i As Long, cod As String, monto As Double, arr As Variant
    If Not IsArray(v) Then Exit Sub
    For i = 1 To UBound(v, 1)
        If MismoFondo(Txt(v(i, cMap("fondo"))), fondo) Then
            cod = SinGuiones(Txt(v(i, cMap("codsbs"))))
            If Len(cod) > 0 Then
                monto = Num(v(i, cMap("valtotal")))
                If Not datos.Exists(cod) Then
                    datos(cod) = Array(Txt(v(i, cMap("nemonico"))), Txt(v(i, cMap("emisor"))), _
                                       Txt(v(i, cMap("moneda"))), 0#, 0#, _
                                       FechaSegura(v(i, cMap("vencimiento"))))
                End If
                arr = datos(cod)
                arr(2 + cual) = Num(arr(2 + cual)) + monto
                datos(cod) = arr
            End If
        End If
    Next i
End Sub


'==============================================================================
'  LEER LAS BASES  ->  Precios, Instrumentos, Marcas, Datos_Fondos
'==============================================================================
Private Sub LeerBases(wsCfg As Worksheet, ByRef nBonos As Long, ByRef nFondos As Long, _
                      ByRef nPrecios As Long, ByRef nMarcas As Long, ByRef avisos As String)
    Dim rutaRF As String, rutaMa As String
    Dim wb As Workbook, ws As Worksheet, yaAbierto As Boolean
    Dim vMae As Variant, vPre As Variant, vMaf As Variant, vDat As Variant, vMar As Variant

    rutaRF = Trim$(Txt(wsCfg.Range("C2").Value))
    rutaMa = Trim$(Txt(wsCfg.Range("C3").Value))
    If Len(rutaRF) = 0 Then Err.Raise vbObjectError + 40, , "Falta la ruta de Base_RF en Config C2."
    If Len(Dir(rutaRF)) = 0 Then Err.Raise vbObjectError + 41, , "No existe Base_RF en: " & rutaRF

    ' ---------- Base_RF
    Paso "Abriendo Base_RF en solo lectura"
    yaAbierto = LibroYaAbierto(rutaRF)
    Set wb = Workbooks.Open(Filename:=rutaRF, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    Set ws = HojaDe(wb, "Maestro")
    If ws Is Nothing Then
        If Not yaAbierto Then wb.Close SaveChanges:=False
        Err.Raise vbObjectError + 42, , "Base_RF no tiene la hoja Maestro."
    End If
    vMae = LeerHoja(ws, 14)
    Set ws = HojaDe(wb, "BD_Precios")
    If ws Is Nothing Then
        If Not yaAbierto Then wb.Close SaveChanges:=False
        Err.Raise vbObjectError + 43, , "Base_RF no tiene la hoja BD_Precios. Corre Cargar vectores alla."
    End If
    vPre = LeerHoja(ws, PR_COLS)
    If Not yaAbierto Then wb.Close SaveChanges:=False
    Set wb = Nothing: Set ws = Nothing

    ' ---------- Base_Marcas (opcional: sin ella no hay fondos, pero hay cuadro)
    If Len(rutaMa) > 0 Then
        If Len(Dir(rutaMa)) = 0 Then Err.Raise vbObjectError + 44, , "No existe Base_Marcas en: " & rutaMa
        Paso "Abriendo Base_Marcas en solo lectura"
        yaAbierto = LibroYaAbierto(rutaMa)
        Set wb = Workbooks.Open(Filename:=rutaMa, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
        Set ws = HojaDe(wb, "Maestro_Fondos")
        If Not ws Is Nothing Then vMaf = LeerHoja(ws, 6)
        Set ws = HojaDe(wb, "Datos_Fondos")
        If Not ws Is Nothing Then vDat = LeerHoja(ws, 7)
        Set ws = HojaDe(wb, "Marcas_Detalle")
        If Not ws Is Nothing Then vMar = LeerHoja(ws, 5)
        If Not yaAbierto Then wb.Close SaveChanges:=False
        Set wb = Nothing: Set ws = Nothing
        If IsEmpty(vMaf) Then avisos = avisos & "Base_Marcas no tiene Maestro_Fondos: sin fondos." & vbCrLf
        If IsEmpty(vMar) Then avisos = avisos & "Base_Marcas no tiene Marcas_Detalle: corre Cargar marcas alla." & vbCrLf
    Else
        avisos = avisos & "Config C3 vacio: el cuadro sale sin fondos." & vbCrLf
    End If

    Paso "Copiando precios"
    nPrecios = CopiarPrecios(vPre)
    Paso "Armando Instrumentos"
    ArmarInstrumentos vMae, vMaf, vDat, CorteEfectivo(wsCfg), nBonos, nFondos, avisos
    Paso "Copiando marcas"
    nMarcas = CopiarMarcas(vMar)
    CopiarDatosFondos vDat
End Sub


Private Function HojaDe(wb As Workbook, nombre As String) As Worksheet
    On Error Resume Next
    Set HojaDe = wb.Worksheets(nombre)
    On Error GoTo 0
End Function


Private Function CopiarPrecios(vPre As Variant) As Long
    ' Hoja Precios: copia tal cual de BD_Precios de la base.
    Dim ws As Worksheet, n As Long, out() As Variant, i As Long, j As Long
    Set ws = HojaOCrea(SH_PRE)
    ws.Cells.Clear
    If Not IsArray(vPre) Then Exit Function
    n = UBound(vPre, 1)
    If n < 1 Then Exit Function
    ReDim out(1 To n, 1 To PR_COLS)
    For i = 1 To n
        For j = 1 To PR_COLS
            out(i, j) = vPre(i, j)
        Next j
        If i > 1 Then out(i, 2) = "'" & SinGuiones(Txt(vPre(i, 2)))
    Next i
    ws.Range(ws.Cells(1, 1), ws.Cells(n, PR_COLS)).Value = out
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, PR_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
    End With
    If n > 1 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(n, 1)).NumberFormat = "dd/mm/yyyy"
        ws.Range(ws.Cells(2, 8), ws.Cells(n, 8)).NumberFormat = "dd/mm/yyyy"
        ws.Range(ws.Cells(2, 5), ws.Cells(n, 7)).NumberFormat = "0.000000"
    End If
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(10).ColumnWidth = 42
    ws.Range("L1").Value = "Copia de BD_Precios de Base_RF. Se rehace en cada ARMAR CUADRO. No la edites."
    ws.Range("L1").Font.Color = RGB(128, 128, 128)
    CongelarPaneles ws, 1, 2
    CopiarPrecios = n - 1
End Function


Private Sub ArmarInstrumentos(vMae As Variant, vMaf As Variant, vDat As Variant, dCorte As Double, _
                              ByRef nBonos As Long, ByRef nFondos As Long, ByRef avisos As String)
    ' Junta el Maestro de Base_RF (bonos) y el de Base_Marcas (fondos) en una
    ' sola hoja, con el nombre final ya armado. Para cada fondo toma de
    ' Datos_Fondos la fila con la ULTIMA fecha <= corte.
    Dim ws As Worksheet, r As Long, i As Long, cod As String, nom As String
    Dim datos As Object, arr As Variant, dFe As Double, vistos As Object
    Dim cab As Variant, monFms As Object, mon As String

    Set ws = HojaOCrea(SH_INS)
    ws.Cells.Clear
    cab = Array("Codigo SBS", "Nombre", "Categoria", "Fuera / REF", "Moneda", "CY (fondos)", _
                "ISIN", "Rating", "Duracion", "Spread (pb)", "Origen", "Nombre base", _
                "Tasa cupon", "Vencimiento", "Fecha datos")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, IN_COLS)).Value = cab
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, IN_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
    End With
    Set vistos = NuevoDic()
    r = 1

    ' --- bonos: Maestro de Base_RF
    '   A cod  B nombre  C cat  D fuera  E nem  F emisor  G cupon  H venc  I mon
    '   J ISIN  K rating  L dur  M spread  N visto
    If IsArray(vMae) Then
        For i = 2 To UBound(vMae, 1)
            cod = SinGuiones(Txt(vMae(i, 1)))
            If Len(cod) > 0 Then
                If vistos.Exists(cod) Then
                    avisos = avisos & "Codigo repetido en Base_RF.Maestro: " & cod & vbCrLf
                ElseIf Len(Trim$(Txt(vMae(i, 3)))) = 0 Then
                    avisos = avisos & "Sin categoria en Base_RF.Maestro (no entra): " & cod & vbCrLf
                Else
                    vistos(cod) = 1
                    r = r + 1
                    nBonos = nBonos + 1
                    nom = Trim$(Txt(vMae(i, 2)))
                    If Len(nom) = 0 Then nom = Trim$(Txt(vMae(i, 5)))
                    If Len(nom) = 0 Then nom = cod
                    ws.Cells(r, 1).Value = "'" & cod
                    ws.Cells(r, 2).Value = ArmaNombre(nom, Num(vMae(i, 7)), Num(FechaSegura(vMae(i, 8))))
                    ws.Cells(r, 3).Value = Trim$(Txt(vMae(i, 3)))
                    ws.Cells(r, 4).Value = UCase$(Trim$(Txt(vMae(i, 4))))
                    ws.Cells(r, 5).Value = NormMoneda(Txt(vMae(i, 9)))
                    If Len(Trim$(Txt(vMae(i, 10)))) > 0 Then ws.Cells(r, 7).Value = "'" & Trim$(Txt(vMae(i, 10)))
                    If Len(Trim$(Txt(vMae(i, 11)))) > 0 Then ws.Cells(r, 8).Value = Trim$(Txt(vMae(i, 11)))
                    If EsNumero(vMae(i, 12)) Then ws.Cells(r, 9).Value = CDbl(vMae(i, 12))
                    If EsNumero(vMae(i, 13)) Then ws.Cells(r, 10).Value = CDbl(vMae(i, 13))
                    ws.Cells(r, 11).Value = "Base_RF"
                    ws.Cells(r, 12).Value = nom
                    If EsNumero(vMae(i, 7)) Then ws.Cells(r, 13).Value = CDbl(vMae(i, 7))
                    If Not IsEmpty(FechaSegura(vMae(i, 8))) Then ws.Cells(r, 14).Value = FechaSegura(vMae(i, 8))
                End If
            End If
        Next i
    End If

    ' --- moneda de los fondos segun el FMS (BD_Pesos: A codigo, D moneda)
    Set monFms = NuevoDic()
    If HojaExiste(SH_PES) Then
        With ThisWorkbook.Worksheets(SH_PES)
            For i = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
                cod = SinGuiones(Txt(.Cells(i, 1).Value))
                If Len(cod) > 0 And Not monFms.Exists(cod) Then
                    If Len(Trim$(Txt(.Cells(i, 4).Value))) > 0 Then monFms(cod) = NormMoneda(Txt(.Cells(i, 4).Value))
                End If
            Next i
        End With
    End If

    ' --- fondos: Maestro_Fondos + Datos_Fondos (ultima fecha <= corte)
    '   Maestro_Fondos: A cod  B nombre  C cat  D fuera  E moneda (opcional)  F ultima marca
    '   Moneda: 1) columna E del Maestro_Fondos si esta llena, 2) FMS (BD_Pesos), 3) PEN
    '   Datos_Fondos:   A fecha  B cod  C nombre  D dur  E spread  F CY  G fuente
    Set datos = NuevoDic()
    If IsArray(vDat) Then
        For i = 2 To UBound(vDat, 1)
            cod = SinGuiones(Txt(vDat(i, 2)))
            If Len(cod) > 0 And EsNumero(vDat(i, 1)) Then
                dFe = CDbl(vDat(i, 1))
                If dFe <= dCorte Then
                    If datos.Exists(cod) Then
                        arr = datos(cod)
                        If dFe >= CDbl(arr(0)) Then datos(cod) = Array(dFe, vDat(i, 4), vDat(i, 5), vDat(i, 6))
                    Else
                        datos(cod) = Array(dFe, vDat(i, 4), vDat(i, 5), vDat(i, 6))
                    End If
                End If
            End If
        Next i
    End If
    If IsArray(vMaf) Then
        For i = 2 To UBound(vMaf, 1)
            cod = SinGuiones(Txt(vMaf(i, 1)))
            If Len(cod) > 0 Then
                If vistos.Exists(cod) Then
                    avisos = avisos & "Codigo repetido (esta en las dos bases): " & cod & vbCrLf
                ElseIf Len(Trim$(Txt(vMaf(i, 3)))) = 0 Then
                    avisos = avisos & "Sin categoria en Base_Marcas.Maestro_Fondos (no entra): " & cod & vbCrLf
                Else
                    vistos(cod) = 1
                    r = r + 1
                    nFondos = nFondos + 1
                    nom = Trim$(Txt(vMaf(i, 2)))
                    If Len(nom) = 0 Then nom = cod
                    ws.Cells(r, 1).Value = "'" & cod
                    ws.Cells(r, 2).Value = nom
                    ws.Cells(r, 3).Value = Trim$(Txt(vMaf(i, 3)))
                    ws.Cells(r, 4).Value = UCase$(Trim$(Txt(vMaf(i, 4))))
                    mon = ""
                    If UBound(vMaf, 2) >= 5 Then mon = Trim$(Txt(vMaf(i, 5)))
                    If Len(mon) > 0 Then
                        mon = NormMoneda(mon)
                    ElseIf monFms.Exists(cod) Then
                        mon = monFms(cod)
                    End If
                    If Len(mon) = 0 Then
                        mon = "PEN"
                        avisos = avisos & "Sin moneda (no esta en el FMS ni en Maestro_Fondos col. E): se asume PEN: " & nom & vbCrLf
                    End If
                    ws.Cells(r, 5).Value = mon
                    ws.Cells(r, 11).Value = "Base_Marcas"
                    ws.Cells(r, 12).Value = nom
                    If datos.Exists(cod) Then
                        arr = datos(cod)
                        If EsNumero(arr(1)) Then ws.Cells(r, 9).Value = CDbl(arr(1))
                        If EsNumero(arr(2)) Then ws.Cells(r, 10).Value = CDbl(arr(2))
                        If EsNumero(arr(3)) Then ws.Cells(r, 6).Value = CDbl(arr(3))
                        ws.Cells(r, 15).Value = CDbl(arr(0))
                    Else
                        avisos = avisos & "Sin datos (dur/spread/CY) a la fecha en Datos_Fondos: " & nom & vbCrLf
                    End If
                End If
            End If
        Next i
    End If

    ws.Columns(1).ColumnWidth = 18
    ws.Columns(2).ColumnWidth = 36
    ws.Columns(3).ColumnWidth = 22
    ws.Range(ws.Columns(4), ws.Columns(10)).ColumnWidth = 11
    ws.Columns(12).ColumnWidth = 28
    ws.Columns(6).NumberFormat = "0.00%"
    ws.Columns(9).NumberFormat = "0.00"
    ws.Columns(10).NumberFormat = "0"
    ws.Columns(13).NumberFormat = "0.000"
    ws.Columns(14).NumberFormat = "dd/mm/yyyy"
    ws.Columns(15).NumberFormat = "dd/mm/yyyy"
    ws.Range("Q1").Value = "Los dos maestros juntos. Se rehace en cada ARMAR CUADRO. Para cambiar algo, cambialo en la base."
    ws.Range("Q1").Font.Color = RGB(128, 128, 128)
    CongelarPaneles ws, 1, 2
End Sub


Private Function CopiarMarcas(vMar As Variant) As Long
    ' Hoja Marcas: copia de Marcas_Detalle (A cod, B fecha, C var, D nombre,
    ' E flag) + LN(1+var) + las seis banderas de ventana, que las pone
    ' ArmarFondos.
    Dim ws As Worksheet, n As Long, out() As Variant, i As Long, k As Long
    Set ws = HojaOCrea(SH_MDE)
    ws.Cells.Clear
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 12)).Value = _
        Array("Codigo SBS", "Fecha", "Var Ajustada", "Nombre en Marcas", "Flag", "LN(1+var)", _
              "WTD", "5D", "MTD", "MayoTD", "YTD", "FY")
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 12))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
    End With
    ws.Columns(1).ColumnWidth = 18
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 14
    ws.Columns(4).ColumnWidth = 30
    ws.Columns(5).ColumnWidth = 30
    ws.Columns(6).ColumnWidth = 14
    ws.Range(ws.Columns(7), ws.Columns(12)).ColumnWidth = 8
    ws.Range("N1").Value = "COMO AUDITAR UN FONDO DEL CUADRO"
    ws.Range("N1").Font.Bold = True
    ws.Range("N2").Value = "1. Filtra por el Codigo SBS.   2. Filtra la ventana (1 en su columna)."
    ws.Range("N3").Value = "3. Retorno = EXP(SUMA de LN(1+var)) - 1.  Es lo que hace la hoja Fondos."
    ws.Range("N1:N3").Font.Color = RGB(100, 100, 100)
    ws.Columns(14).ColumnWidth = 70
    CongelarPaneles ws, 1, 2
    If Not IsArray(vMar) Then Exit Function
    n = UBound(vMar, 1)
    If n < 2 Then Exit Function
    ReDim out(1 To n - 1, 1 To 6)
    k = 0
    For i = 2 To n
        If EsNumero(vMar(i, 2)) And EsNumero(vMar(i, 3)) Then
            k = k + 1
            out(k, 1) = "'" & SinGuiones(Txt(vMar(i, 1)))
            out(k, 2) = CDbl(vMar(i, 2))
            out(k, 3) = CDbl(vMar(i, 3))
            out(k, 4) = Txt(vMar(i, 4))
            out(k, 5) = Txt(vMar(i, 5))
            If CDbl(vMar(i, 3)) > -1 Then out(k, 6) = Log(1 + CDbl(vMar(i, 3)))
        End If
    Next i
    If k = 0 Then Exit Function
    ws.Range(ws.Cells(2, 1), ws.Cells(k + 1, 6)).Value = out
    ws.Range(ws.Cells(2, 2), ws.Cells(k + 1, 2)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 3), ws.Cells(k + 1, 3)).NumberFormat = "0.0000%"
    ws.Range(ws.Cells(2, 6), ws.Cells(k + 1, 6)).NumberFormat = "0.00000000"
    CopiarMarcas = k
End Function


Private Sub CopiarDatosFondos(vDat As Variant)
    Dim ws As Worksheet, n As Long, out() As Variant, i As Long, j As Long
    Set ws = HojaOCrea(SH_DAT)
    ws.Cells.Clear
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 7)).Value = _
        Array("Fecha", "Codigo SBS", "Nombre", "Duracion", "Spread (pb)", "CY", "Fuente")
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 7))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(64, 64, 64)
    End With
    ws.Range("I1").Value = "Copia del historico de Base_Marcas. El cuadro usa, por fondo, la fila con la ultima fecha <= corte."
    ws.Range("I1").Font.Color = RGB(128, 128, 128)
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 30
    ws.Columns(7).ColumnWidth = 28
    If Not IsArray(vDat) Then Exit Sub
    n = UBound(vDat, 1)
    If n < 2 Then Exit Sub
    ReDim out(1 To n - 1, 1 To 7)
    For i = 2 To n
        For j = 1 To 7
            out(i - 1, j) = vDat(i, j)
        Next j
        out(i - 1, 2) = "'" & SinGuiones(Txt(vDat(i, 2)))
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(n, 7)).Value = out
    ws.Range(ws.Cells(2, 1), ws.Cells(n, 1)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 4), ws.Cells(n, 4)).NumberFormat = "0.00"
    ws.Range(ws.Cells(2, 5), ws.Cells(n, 5)).NumberFormat = "0"
    ws.Range(ws.Cells(2, 6), ws.Cells(n, 6)).NumberFormat = "0.00%"
End Sub


Private Function LeerInstrumentos(ByRef orden As Variant) As Object
    ' Lee la hoja Instrumentos (ya armada) y devuelve, por codigo:
    '   Array(0 codigo, 1 nombre, 2 categoria, 3 monedaBloque, 4 fuera, 5 cy,
    '         6 vintage(no se usa), 7 monedaReal, 8 esRef, 9 ISIN, 10 dur,
    '         11 spread, 12 rating)
    Dim ws As Worksheet, d As Object, i As Long, ult As Long
    Dim cod As String, cat As String, mon As String, fue As String
    Dim lst() As String, n As Long

    Set d = NuevoDic()
    orden = Array()
    Set LeerInstrumentos = d
    If Not HojaExiste(SH_INS) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_INS)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function

    ReDim lst(0 To ult)
    n = -1
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
        cat = Trim$(Txt(ws.Cells(i, 3).Value))
        If Len(cod) > 0 And Len(cat) > 0 And Not d.Exists(cod) Then
            mon = NormMoneda(Txt(ws.Cells(i, 5).Value))
            fue = UCase$(Trim$(Txt(ws.Cells(i, 4).Value)))
            d(cod) = Array(cod, Trim$(Txt(ws.Cells(i, 2).Value)), cat, MonedaBloque(cat, mon), _
                           (Left$(fue & "  ", 2) = "SI"), ws.Cells(i, 6).Value, 0#, mon, _
                           (Left$(fue & "   ", 3) = "REF"), Trim$(Txt(ws.Cells(i, 7).Value)), _
                           ws.Cells(i, 9).Value, ws.Cells(i, 10).Value, Trim$(Txt(ws.Cells(i, 8).Value)))
            n = n + 1
            lst(n) = cod
        End If
    Next i
    If n >= 0 Then
        ReDim Preserve lst(0 To n)
        orden = lst
    End If
End Function


'==============================================================================
'  FONDOS: retorno compuesto por ventana, desde la hoja Marcas
'==============================================================================
Private Sub ArmarFondos(ins As Object, orden As Variant, dCorte As Double, wsCfg As Worksheet, _
                        ByRef nConMarca As Long, ByRef nUsadas As Long)
    Dim wsMar As Worksheet, wsFon As Worksheet, wsIns As Worksheet
    Dim ult As Long, i As Long, j As Long, fila As Long, nDet As Long
    Dim v As Variant, bases() As Double, dTope As Double, usarPost As Boolean
    Dim cod As String, dFe As Double, ln As Double
    Dim acum As Object, ultMar As Object, arr As Variant, k As Variant, arrIns As Variant
    Dim flags() As Variant, cyIns As Object

    bases = BasesVentana(wsCfg, dCorte)
    usarPost = (UCase$(Txt(wsCfg.Range("C14").Value)) <> "NO")
    If usarPost Then dTope = 2958465# Else dTope = dCorte

    Set acum = NuevoDic()
    Set ultMar = NuevoDic()
    Set wsMar = HojaOCrea(SH_MDE)
    ult = wsMar.Cells(wsMar.Rows.Count, 1).End(xlUp).Row
    nDet = IIf(ult > 1, ult - 1, 0)

    If nDet > 0 Then
        v = wsMar.Range(wsMar.Cells(2, 1), wsMar.Cells(ult, 6)).Value2
        ReDim flags(1 To nDet, 1 To N_VENT)
        For i = 1 To nDet
            cod = SinGuiones(Txt(v(i, 1)))
            If ins.Exists(cod) And EsNumero(v(i, 2)) And EsNumero(v(i, 6)) Then
                dFe = CDbl(v(i, 2))
                ln = CDbl(v(i, 6))
                If Not ultMar.Exists(cod) Then ultMar(cod) = 0
                If dFe <= dTope Then
                    If dFe > CDbl(ultMar(cod)) Then ultMar(cod) = dFe
                    If Not acum.Exists(cod) Then acum(cod) = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                    arr = acum(cod)
                    For j = 0 To N_VENT - 1
                        If dFe > bases(j) Then
                            arr(j) = CDbl(arr(j)) + ln
                            arr(6 + j) = CLng(arr(6 + j)) + 1
                            flags(i, 1 + j) = 1
                        End If
                    Next j
                    acum(cod) = arr
                    nUsadas = nUsadas + 1
                End If
            End If
        Next i
        wsMar.Range(wsMar.Cells(2, 7), wsMar.Cells(ult, 12)).Value = flags
    End If

    ' CY de los fondos: columna F de Instrumentos (viene de Datos_Fondos)
    Set cyIns = NuevoDic()
    Set wsIns = ThisWorkbook.Worksheets(SH_INS)
    ult = wsIns.Cells(wsIns.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuiones(Txt(wsIns.Cells(i, 1).Value))
        If Len(cod) > 0 Then cyIns(cod) = wsIns.Cells(i, 6).Value
    Next i

    Set wsFon = HojaFondos()
    fila = 1
    For Each k In acum.Keys
        fila = fila + 1
        cod = CStr(k)
        arr = acum(k)
        wsFon.Cells(fila, 1).Value = "'" & cod
        If ins.Exists(cod) Then
            arrIns = ins(cod)
            wsFon.Cells(fila, 2).Value = CStr(arrIns(1))
        End If
        If cyIns.Exists(cod) Then
            If EsNumero(cyIns(cod)) Then wsFon.Cells(fila, 3).Value = CDbl(cyIns(cod))
        End If
        ' D a I: el numero (lo lee el cuadro). L a Q: la misma cuenta como
        ' formula sobre Marcas, para auditar. R a W: cuantas marcas entraron.
        For j = 0 To N_VENT - 1
            If CLng(arr(6 + j)) > 0 Then wsFon.Cells(fila, 4 + j).Value = Exp(CDbl(arr(j))) - 1
            wsFon.Cells(fila, 12 + j).Formula = ChequeoVentana(fila, j, nDet)
            wsFon.Cells(fila, 18 + j).Value = CLng(arr(6 + j))
        Next j
        wsFon.Cells(fila, 10).Value = CDbl(ultMar(k))
        wsFon.Cells(fila, 10).NumberFormat = "dd/mm/yyyy"
        nConMarca = nConMarca + 1
    Next k
    If fila > 1 Then
        wsFon.Range(wsFon.Cells(2, 3), wsFon.Cells(fila, 9)).NumberFormat = "0.00%"
        wsFon.Range(wsFon.Cells(2, 12), wsFon.Cells(fila, 17)).NumberFormat = "0.00%"
    End If
End Sub


Private Function HojaFondos() As Worksheet
    Dim ws As Worksheet, cab As Variant
    Set ws = HojaOCrea(SH_FON)
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone
    cab = Array("Codigo SBS", "Nombre", "Curr. Yield", _
                "WTD", "5D", "MTD", "MayoTD", "YTD", "FY", "Ultima marca", "", _
                "Chequeo WTD", "Chequeo 5D", "Chequeo MTD", "Chequeo MayoTD", "Chequeo YTD", "Chequeo FY", _
                "n WTD", "n 5D", "n MTD", "n MayoTD", "n YTD", "n FY")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 23)).Value = cab
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 10))
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
    End With
    With ws.Range(ws.Cells(1, 12), ws.Cells(1, 23))
        .Font.Bold = True
        .Interior.Color = RGB(127, 127, 127)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("Y1").Value = "L a Q: la misma cuenta, como formula sobre la hoja Marcas. Si D:I y L:Q coinciden, esta bien."
    ws.Range("Y2").Value = "R a W: cuantas marcas entraron en cada ventana. n = 0 -> celda vacia a proposito."
    ws.Range("Y3").Value = "C: Curr. Yield del fondo, de Datos_Fondos (ultima fecha <= corte)."
    ws.Range("Y1:Y3").Font.Color = RGB(100, 100, 100)
    ws.Columns(2).ColumnWidth = 36
    ws.Range(ws.Columns(12), ws.Columns(23)).ColumnWidth = 12
    ws.Columns(25).ColumnWidth = 80
    ws.Range("A2:A500").NumberFormat = "@"
    CongelarPaneles ws, 1, 2
    Set HojaFondos = ws
End Function


Private Function ChequeoVentana(fila As Long, j As Long, nDet As Long) As String
    ' Apunta a la hoja Marcas DIRECTO, no a nombres definidos.
    Dim rCod As String, rLn As String, rW As String, f As String
    If nDet < 1 Then Exit Function
    rCod = RefMarcas(1, nDet)
    rLn = RefMarcas(6, nDet)
    rW = RefMarcas(7 + j, nDet)
    f = "=IFERROR(IF(COUNTIFS(" & rCod & ",$A" & fila & "," & rW & ",1)=0,"""","
    f = f & "EXP(SUMIFS(" & rLn & "," & rCod & ",$A" & fila & "," & rW & ",1))-1),""ERROR"")"
    ChequeoVentana = f
End Function


Private Function RefMarcas(col As Long, nDet As Long) As String
    Dim L As String
    L = NumALetra(col)
    RefMarcas = SH_MDE & "!$" & L & "$2:$" & L & "$" & (nDet + 1)
End Function


'==============================================================================
'  ARMAR CUADRO  (la macro del dia)
'==============================================================================
Public Sub ArmarTodo()
    Dim wsCfg As Worksheet, dCorte As Double
    Dim ins As Object, orden As Variant
    Dim nBonos As Long, nFondos As Long, nPrecios As Long, nMarcas As Long
    Dim nConMarca As Long, nUsadas As Long, avisos As String
    Dim scrPrev As Boolean, calcPrev As XlCalculation, evPrev As Boolean, alertPrev As Boolean

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

    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    If Not HojaExiste(SH_PES) Then Err.Raise vbObjectError + 20, , _
        "No hay BD_Pesos. Corre Cargar FMS primero."

    LeerBases wsCfg, nBonos, nFondos, nPrecios, nMarcas, avisos
    If nPrecios = 0 Then Err.Raise vbObjectError + 21, , _
        "BD_Precios de Base_RF esta vacia. Corre Cargar vectores alla."

    dCorte = CorteEfectivo(wsCfg)
    Paso "Corte: " & Format$(CDate(dCorte), "dd/mm/yyyy")

    Paso "Leyendo Instrumentos"
    Set ins = LeerInstrumentos(orden)
    If ins.Count = 0 Then Err.Raise vbObjectError + 22, , _
        "Ningun instrumento con categoria. Revisa los maestros de las bases."

    Paso "Armando Retornos (bonos)"
    ArmarRetornos ins, orden, dCorte, wsCfg

    Paso "Armando Fondos"
    ArmarFondos ins, orden, dCorte, wsCfg, nConMarca, nUsadas

    Paso "Armando el CUADRO"
    ArmarCuadro ins, orden, dCorte
    OcultarHojasAux

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    Restaurar calcPrev, scrPrev, evPrev
    Application.DisplayAlerts = alertPrev

    Aviso "v" & VER & " - ARMAR CUADRO terminado en " & Format$(Timer - mT0, "0.0") & " s." & vbCrLf & vbCrLf & _
          mPasos & vbCrLf & _
          "Corte: " & Format$(CDate(dCorte), "dd/mm/yyyy") & vbCrLf & _
          "Bonos: " & nBonos & "   Fondos: " & nFondos & " (" & nConMarca & " con marcas, " & _
          Format$(nUsadas, "#,##0") & " marcas usadas)" & vbCrLf & _
          "Precios copiados: " & Format$(nPrecios, "#,##0") & "   Marcas copiadas: " & Format$(nMarcas, "#,##0") & vbCrLf & vbCrLf & _
          DiagMontos() & DiagOrden() & _
          IIf(Len(avisos) > 0, vbCrLf & "AVISOS:" & vbCrLf & avisos, "") & vbCrLf & _
          "Revisa la columna Flag de Retornos y la hoja Fondos.", "ArmarTodo"
    Exit Sub

Falla:
    Application.DisplayAlerts = True
    Reventar Nothing, calcPrev, scrPrev, evPrev, "ArmarTodo"
End Sub


Private Function DiagOrden() As String
    DiagOrden = "Orden: por peso F1 de mayor a menor; si F1 empata, por F2." & vbCrLf
End Function


Public Sub RevisarCodigosFMS()
    ' Que codigos tienen peso en el FMS y NO estan en ningun maestro.
    ' Son posiciones del fondo que el cuadro no muestra.
    Dim wsPes As Worksheet, wsIns As Worksheet, i As Long, ult As Long, cod As String
    Dim ins As Object, msg As String, n As Long, w As Double

    On Error GoTo Falla
    If Not HojaExiste(SH_PES) Then
        MsgBox "No hay BD_Pesos. Corre Cargar FMS.", vbExclamation, "RevisarCodigosFMS"
        Exit Sub
    End If
    If Not HojaExiste(SH_INS) Then
        MsgBox "No hay Instrumentos. Corre ARMAR CUADRO una vez.", vbExclamation, "RevisarCodigosFMS"
        Exit Sub
    End If
    Set ins = NuevoDic()
    Set wsIns = ThisWorkbook.Worksheets(SH_INS)
    ult = wsIns.Cells(wsIns.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuiones(Txt(wsIns.Cells(i, 1).Value))
        If Len(cod) > 0 Then ins(cod) = 1
    Next i

    Set wsPes = ThisWorkbook.Worksheets(SH_PES)
    ult = wsPes.Cells(wsPes.Rows.Count, 1).End(xlUp).Row
    msg = "Codigos con peso en el FMS que NO estan en los maestros:" & vbCrLf & vbCrLf
    For i = 2 To ult
        cod = SinGuiones(Txt(wsPes.Cells(i, 1).Value))
        w = Num(wsPes.Cells(i, 7).Value) + Num(wsPes.Cells(i, 8).Value)
        If Len(cod) > 0 And Not ins.Exists(cod) And w >= 0.0005 Then
            n = n + 1
            If n <= 40 Then
                msg = msg & cod & "  " & Left$(Txt(wsPes.Cells(i, 2).Value), 22) & _
                      "  F1 " & Format$(Num(wsPes.Cells(i, 7).Value), "0.00%") & _
                      "  F2 " & Format$(Num(wsPes.Cells(i, 8).Value), "0.00%") & vbCrLf
            End If
        End If
    Next i
    If n = 0 Then
        msg = "Todos los codigos con peso >= 0.05% estan en los maestros."
    Else
        msg = msg & vbCrLf & n & " codigos (se muestran hasta 40, peso >= 0.05%)." & vbCrLf & _
              "Si alguno debe entrar al cuadro, agregalo al Maestro de la base que corresponda."
    End If
    MsgBox msg, vbInformation, "RevisarCodigosFMS"
    Exit Sub

Falla:
    MsgBox "RevisarCodigosFMS se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "RevisarCodigosFMS"
End Sub


'==============================================================================
'  DIAGNOSTICO, PANEL, VERSION
'==============================================================================
Public Sub Diagnostico()
    Dim wsCfg As Worksheet, ws As Worksheet, msg As String, ult As Long, r As String

    On Error GoTo Falla
    msg = "PRICE ACTION  v" & VER & vbCrLf & String$(44, "-") & vbCrLf
    If Not HojaExiste(SH_CFG) Then
        MsgBox msg & "No existe Config. Corre Crear Config.", vbExclamation, "Diagnostico"
        Exit Sub
    End If
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    r = Trim$(Txt(wsCfg.Range("C2").Value))
    msg = msg & "Base_RF: " & IIf(Len(r) = 0, "VACIO (C2)", IIf(Len(Dir(r)) > 0, "ok", "NO EXISTE: " & r)) & vbCrLf
    r = Trim$(Txt(wsCfg.Range("C3").Value))
    msg = msg & "Base_Marcas: " & IIf(Len(r) = 0, "VACIO (C3) - sin fondos", IIf(Len(Dir(r)) > 0, "ok", "NO EXISTE: " & r)) & vbCrLf
    r = SinBarra(Txt(wsCfg.Range("C4").Value))
    msg = msg & "Carpeta FMS: " & IIf(Len(r) = 0, "VACIO (C4)", r) & vbCrLf
    msg = msg & "Corte: " & IIf(IsDate(wsCfg.Range("C5").Value), Txt(wsCfg.Range("C5").Text), "vacio = ultimo precio de la base") & vbCrLf & vbCrLf

    If HojaExiste(SH_PES) Then
        Set ws = ThisWorkbook.Worksheets(SH_PES)
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        msg = msg & "BD_PESOS: " & IIf(ult > 1, ult - 1, 0) & " codigos"
        If ult > 1 Then msg = msg & ", FMS del " & Txt(ws.Cells(2, 9).Text)
        msg = msg & vbCrLf & DiagMontos()
    Else
        msg = msg & "BD_PESOS: no existe. Corre Cargar FMS." & vbCrLf
    End If
    msg = msg & vbCrLf
    msg = msg & "Precios: " & FilasDe(SH_PRE) & "   Instrumentos: " & FilasDe(SH_INS) & _
          "   Marcas: " & FilasDe(SH_MDE) & "   Datos_Fondos: " & FilasDe(SH_DAT) & vbCrLf
    msg = msg & "Retornos: " & FilasDe(SH_RET) & "   Fondos: " & FilasDe(SH_FON) & _
          "   Manual: " & FilasDe(SH_MAN) & vbCrLf
    msg = msg & "(filas de datos; 0 = la hoja no existe o esta vacia)"
    MsgBox msg, vbInformation, "Diagnostico"
    Exit Sub

Falla:
    MsgBox "Diagnostico se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Diagnostico"
End Sub


Private Function FilasDe(hoja As String) As Long
    Dim ws As Worksheet, u As Long
    If Not HojaExiste(hoja) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(hoja)
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If u > 1 Then FilasDe = u - 1
End Function


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

    ws.Range("B2").Value = "Retornos: Cuasi Sob & Iliquidos"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 16
    ws.Range("B2").Font.Color = RGB(192, 0, 0)
    ws.Range("B3").Value = "v" & VER & "  -  lee Base_RF, Base_Marcas y el FMS. No escribe en las bases."
    ws.Range("B3").Font.Color = RGB(128, 128, 128)

    ws.Range("B5").Value = "RUTINA DEL DIA"
    ws.Range("B5").Font.Bold = True
    ws.Range("F5").Value = "CUANDO HACE FALTA"
    ws.Range("F5").Font.Bold = True

    rutina = Array( _
        "CargarFMS", "1. Cargar FMS", _
        "ArmarTodo", "2. ARMAR CUADRO", _
        "GuardarCuadro", "3. Guardar la foto")

    extra = Array( _
        "NormalizarExcel", "ARREGLAR EXCEL", _
        "CrearConfig", "Crear Config", _
        "CrearManual", "Crear Manual", _
        "CargarFMSForzado", "Cargar FMS FORZADO", _
        "ProbarTipoCambio", "Probar tipo de cambio", _
        "RevisarCodigosFMS", "Codigos del FMS sin seguir", _
        "MostrarHojas", "Mostrar hojas (auditar)", _
        "OcultarHojas", "Ocultar hojas", _
        "Diagnostico", "Diagnostico", _
        "Version", "Version")

    BotonesEnColumna ws, rutina, 60, 110, 170, 30, True
    BotonesEnColumna ws, extra, 300, 110, 170, 24, False

    ws.Range("B26").Value = "Antes del 1: en Base_RF corre 'Cargar vectores'. Si hay marcas nuevas, en Base_Marcas 'Cargar marcas'."
    ws.Range("B27").Value = "El 2 abre las dos bases en solo lectura, se copia lo que necesita y arma todo. Las bases no se tocan."
    ws.Range("B28").Value = "Si el 1 dice que ya tiene el FMS del dia y lo corrigieron: Cargar FMS FORZADO."
    ws.Range("B29").Value = "El 3 necesita el modulo mod_Guardar."
    ws.Range("B30").Value = "Si cortaste una macro con Esc: ARREGLAR EXCEL antes de seguir."
    ws.Range("B31").Value = "Precios, Instrumentos, Marcas, Retornos y Fondos quedan ocultas: 'Mostrar hojas' para auditar."
    ws.Range("B33").Value = "CUADRO: total return = variacion de precio limpio + cupon devengado (dias/365). Peso = Val_total / total del fondo (FMS)."
    ws.Range("B34").Value = "Filas grises: promedio ponderado por F1. USD pasa a soles con el spread de Config C15 prorrateado por dias."
    ws.Range("B35").Value = "SubTotales y Total: + prima de Config C22 (modo C23). USD mm = peso x total del fondo / tipo de cambio."
    ws.Range("B26:B35").Font.Color = RGB(128, 128, 128)

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


Public Sub Version()
    MsgBox "mod_PriceAction  v" & VER, vbInformation, "Version"
End Sub


Public Sub OcultarHojas()
    OcultarHojasAux
    ThisWorkbook.Worksheets(SH_CUA).Activate
End Sub


Public Sub MostrarHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub


Private Sub OcultarHojasAux()
    ' A la vista solo lo que se usa: Panel, Config, Manual y CUADRO. El resto
    ' son copias de las bases y calculos intermedios: se ocultan, no se borran.
    ' Boton "Mostrar hojas" para auditar.
    Dim ws As Worksheet, nom As String
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        nom = ws.Name
        If nom = SH_PES Or nom = SH_PRE Or nom = SH_INS Or nom = SH_MDE Or nom = SH_DAT _
           Or nom = SH_RET Or nom = SH_FON Or nom = SH_LOG Or nom = SH_MVA Then
            ws.Visible = xlSheetHidden
        End If
    Next ws
    On Error GoTo 0
End Sub


'==============================================================================
'  REUTILIZADO DE mod_Iliquidos (FMS, Retornos, CUADRO, Manual, utilidades)
'==============================================================================
'==============================================================================
'  CUADRO  (sin columnas de ayuda: cada fila gris apunta a filas concretas)
'==============================================================================
Private Sub ArmarCuadro(ins As Object, orden As Variant, dCorte As Double)
    Dim ws As Worksheet, r As Long, ultFila As Long
    Dim monedas As Variant, cats As Variant
    Dim im As Long, ic As Long, i As Long
    Dim arr As Variant, mon As String, cat As String
    Dim vent As Variant, j As Long, bases() As Double
    Dim fondos As Object, lista As Variant, pesos As Object
    Dim man As Variant, manVar As Object
    Dim tipos As Object, filasSub As Object, filasBlq As Object, catDeSub As Object
    Dim rBlq As Long, rSub As Long, a As Long, b As Long
    Dim subsBlq As String, blqs As String
    Dim catsA As Variant, catsB As Variant, etiq As String, hayNiv As Boolean
    Dim hayFuera As Boolean, k As Variant

    TotalesFondo
    man = LeerManual()
    Set manVar = LeerManualVar(dCorte)
    Set fondos = LeerFondos()
    Set pesos = PesosPorCodigo()
    vent = Array("WTD", "5D", "MTD", "MayoTD", "YTD", "FY")
    bases = BasesVentana(ThisWorkbook.Worksheets(SH_CFG), dCorte)
    Set tipos = NuevoDic()       ' fila -> BLQ / SUB / INS / REF / MAN / OUT / ETQ / NIV / TOT / SEP
    Set filasSub = NuevoDic()    ' fila de categoria -> "a:b" (su tramo de instrumentos)
    Set filasBlq = NuevoDic()    ' fila de bloque -> lista de filas de categoria "r1|r2|..."
    Set catDeSub = NuevoDic()    ' fila de categoria -> nombre de la categoria

    Set ws = HojaOCrea(SH_CUA)
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    ' --- fila 1: titulo, tipo de cambio, corte
    ws.Range("A1").Value = "Retornos: Cuasi Sob & Iliquidos"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 10
    ws.Range("A1").Font.Color = RGB(192, 0, 0)
    ws.Cells(1, CC_U1 - 1).Value = "Tipo de cambio"
    ws.Cells(1, CC_U1 - 1).HorizontalAlignment = xlRight
    ws.Cells(1, CC_U1 - 1).Font.Color = RGB(128, 128, 128)
    ws.Cells(1, CC_U1).Formula = "=IFERROR(IF(N(" & SH_CFG & "!$C$24)>0," & SH_CFG & "!$C$24," & SH_PES & "!$M$3),"""")"
    ws.Cells(1, CC_U1).NumberFormat = "0.0000"
    ws.Cells(1, CC_U1).Font.Bold = True
    ws.Cells(1, CC_W1).Value = dCorte
    ws.Cells(1, CC_W1).NumberFormat = "dd/mm/yyyy"
    ws.Cells(1, CC_W1).Font.Bold = True
    ws.Cells(1, CC_W1 + 1).Value = "<- corte (Config C5)"
    ws.Cells(1, CC_W1 + 1).Font.Color = RGB(128, 128, 128)

    ' --- filas 2 y 3: fin y base de cada ventana; fila 4: cabecera
    ws.Cells(CU_FIN, CC_NOM).Value = "fin"
    ws.Cells(CU_BASE, CC_NOM).Value = "base"
    ws.Range(ws.Cells(CU_FIN, CC_NOM), ws.Cells(CU_BASE, CC_NOM)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(CU_FIN, CC_NOM), ws.Cells(CU_BASE, CC_NOM)).Font.Color = RGB(128, 128, 128)
    For j = 0 To N_VENT - 1
        ws.Cells(CU_FIN, CC_W1 + j).Value = dCorte
        ws.Cells(CU_BASE, CC_W1 + j).Value = bases(j)
        ws.Cells(CU_CAB, CC_W1 + j).Value = vent(j)
    Next j
    ws.Cells(CU_CAB, CC_NOM).Value = "Instrumento"
    ws.Cells(CU_CAB, CC_COD).Value = "Codigo SBS"
    ws.Cells(CU_CAB, CC_MRE).Value = "Mon"
    ws.Cells(CU_CAB, CC_ISIN).Value = "ISIN"
    ws.Cells(CU_CAB, CC_RAT).Value = "Rating"
    ws.Cells(CU_CAB, CC_DUR).Value = "Dur"
    ws.Cells(CU_CAB, CC_SPR).Value = "Spread (pb)"
    ws.Cells(CU_CAB, CC_CY).Value = "CY"
    ws.Cells(CU_CAB, CC_CYP).Value = "CY_PEN"
    ws.Cells(CU_CAB, CC_P1).Value = "F1"
    ws.Cells(CU_CAB, CC_P1 + 1).Value = "F2"
    ws.Cells(CU_CAB, CC_U1).Value = "F1"
    ws.Cells(CU_CAB, CC_U1 + 1).Value = "F2"
    ws.Cells(CU_BASE, CC_U1).Value = "USD mm"
    With ws.Range(ws.Cells(CU_BASE, CC_U1), ws.Cells(CU_BASE, CC_U1 + 1))
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 8
        .Font.Color = RGB(80, 80, 80)
    End With
    ws.Cells(CU_CAB, CC_MEN).Value = "Mensual YTD"
    ws.Cells(CU_CAB, CC_ANN).Value = "Ann YTD"

    ' --- bloques por moneda, categorias en el orden de Config C17, instrumentos
    '     por peso F1 (desempate F2). REF al final de su categoria, sin peso.
    monedas = OrdenMonedas(ins, orden)
    r = CU_DAT - 1
    blqs = ""
    For im = LBound(monedas) To UBound(monedas)
        mon = CStr(monedas(im))
        r = r + 1
        rBlq = r
        ws.Cells(r, CC_NOM).Value = mon
        tipos(r) = "BLQ"
        blqs = blqs & IIf(Len(blqs) > 0, "|", "") & r
        subsBlq = ""
        cats = OrdenCategorias(ins, orden, mon)
        For ic = LBound(cats) To UBound(cats)
            cat = CStr(cats(ic))
            r = r + 1
            rSub = r
            ws.Cells(r, CC_NOM).Value = cat
            tipos(r) = "SUB"
            catDeSub(rSub) = cat
            subsBlq = subsBlq & IIf(Len(subsBlq) > 0, "|", "") & rSub
            a = r + 1
            lista = OrdenFilas(ins, orden, mon, cat, pesos)
            If IsArray(lista) Then
                If UBound(lista) >= LBound(lista) Then
                    ' primero los normales, despues las referencias
                    For i = LBound(lista) To UBound(lista)
                        arr = ins(lista(i))
                        If Not EsRef(arr) Then
                            r = r + 1
                            FilaInstrumento ws, r, CStr(arr(0)), CStr(arr(1)), "INS", fondos, arr(5), CStr(arr(7)), arr
                            tipos(r) = "INS"
                            r = FilasManuales(ws, r, man, manVar, CStr(arr(0)), CStr(arr(1)), tipos)
                        End If
                    Next i
                    For i = LBound(lista) To UBound(lista)
                        arr = ins(lista(i))
                        If EsRef(arr) Then
                            r = r + 1
                            FilaInstrumento ws, r, CStr(arr(0)), CStr(arr(1)), "REF", fondos, arr(5), CStr(arr(7)), arr
                            tipos(r) = "REF"
                        End If
                    Next i
                End If
            End If
            b = r
            filasSub(rSub) = a & ":" & b
        Next ic
        filasBlq(rBlq) = subsBlq
    Next im

    ' --- SubTotal Fondos / SubTotal Fondos + Coinversiones (sobre las filas de
    '     categoria de los DOS bloques) y Total (sobre los bloques)
    catsA = CatsDeConfig("C18")
    catsB = CatsDeConfig("C19")
    etiq = EtiquetaNiveles()
    If IsArray(catsA) Then
        r = r + 2
        ws.Cells(r, CC_NOM).Value = "SubTotal Fondos"
        tipos(r) = "ETQ"
        r = r + 1
        ws.Cells(r, CC_NOM).Value = etiq
        tipos(r) = "NIV"
        FilaExplicita ws, r, SubsDeCats(catDeSub, catsA), True
        hayNiv = True
    End If
    If IsArray(catsB) Then
        r = r + 2
        ws.Cells(r, CC_NOM).Value = "SubTotal Fondos + Coinversiones"
        tipos(r) = "ETQ"
        r = r + 1
        ws.Cells(r, CC_NOM).Value = etiq
        tipos(r) = "NIV"
        FilaExplicita ws, r, SubsDeCats(catDeSub, catsB), True
        hayNiv = True
    End If
    If hayNiv Then r = r + 2 Else r = r + 1
    ws.Cells(r, CC_NOM).Value = "Total"
    tipos(r) = "TOT"
    FilaExplicita ws, r, blqs, True

    ' --- fuera del Total
    For i = LBound(orden) To UBound(orden)
        arr = ins(orden(i))
        If arr(4) Then hayFuera = True
    Next i
    If hayFuera Then
        r = r + 2
        ws.Cells(r, CC_NOM).Value = "Fuera del Total"
        ws.Cells(r, CC_NOM).Font.Italic = True
        ws.Cells(r, CC_NOM).Font.Size = 8
        ws.Cells(r, CC_NOM).Font.Color = RGB(128, 128, 128)
        tipos(r) = "SEP"
        For i = LBound(orden) To UBound(orden)
            arr = ins(orden(i))
            If arr(4) Then
                r = r + 1
                FilaInstrumento ws, r, CStr(arr(0)), CStr(arr(1)), "OUT", fondos, arr(5), CStr(arr(7)), arr
                tipos(r) = "OUT"
            End If
        Next i
    End If
    ultFila = r

    ' --- filas grises: categoria = su tramo; bloque = sus categorias
    For Each k In filasSub.Keys
        FilaCategoria ws, CLng(k), CLng(Split(CStr(filasSub(k)), ":")(0)), CLng(Split(CStr(filasSub(k)), ":")(1))
    Next k
    For Each k In filasBlq.Keys
        FilaExplicita ws, CLng(k), CStr(filasBlq(k)), False
    Next k

    FormatearCuadro ws, ultFila, tipos
End Sub


Private Function SubsDeCats(catDeSub As Object, cats As Variant) As String
    ' Filas de categoria (de los dos bloques) cuyo nombre esta en la lista.
    Dim k As Variant, j As Long, out As String
    If Not IsArray(cats) Then Exit Function
    For Each k In catDeSub.Keys
        For j = LBound(cats) To UBound(cats)
            If NormTxt(CStr(catDeSub(k))) = NormTxt(CStr(cats(j))) Then
                out = out & IIf(Len(out) > 0, "|", "") & k
                Exit For
            End If
        Next j
    Next k
    SubsDeCats = out
End Function


Private Sub FilaInstrumento(ws As Worksheet, r As Long, cod As String, nom As String, tipo As String, _
                            fondos As Object, cyMan As Variant, monReal As String, ref As Variant)
    Dim j As Long, cl As String, colYTD As String
    Dim esFondo As Boolean, f As Variant, cCod As String, cMon As String

    If Not fondos Is Nothing Then esFondo = fondos.Exists(cod)
    If esFondo Then f = fondos(cod)

    mDetalle = "CUADRO: " & nom & " en la fila " & r
    ws.Cells(r, CC_NOM).Value = nom
    ws.Cells(r, CC_COD).Value = "'" & cod
    ws.Cells(r, CC_MRE).Value = monReal
    If IsArray(ref) Then
        If UBound(ref) >= 12 Then
            If Len(Txt(ref(9))) > 0 Then ws.Cells(r, CC_ISIN).Value = "'" & Txt(ref(9))
            If EsNumero(ref(10)) Then ws.Cells(r, CC_DUR).Value = CDbl(ref(10))
            If EsNumero(ref(11)) Then ws.Cells(r, CC_SPR).Value = CDbl(ref(11))
            If Len(Txt(ref(12))) > 0 Then ws.Cells(r, CC_RAT).Value = Txt(ref(12))
        End If
    End If
    If tipo = "INS" Or tipo = "REF" Then ws.Cells(r, CC_NOM).IndentLevel = 2

    cCod = "$" & NumALetra(CC_COD) & r
    cMon = "$" & NumALetra(CC_MRE) & r

    ' Peso del FMS. Una REF no lleva peso: asi queda fuera de todo promedio
    ' sin necesitar una columna que diga "no me cuentes".
    If tipo <> "REF" Then
        ws.Cells(r, CC_P1).Formula = "=IFERROR(SUMIFS(BD_Pesos!$G$2:$G$20000,BD_Pesos!$A$2:$A$20000," & cCod & "),0)"
        ws.Cells(r, CC_P1 + 1).Formula = "=IFERROR(SUMIFS(BD_Pesos!$H$2:$H$20000,BD_Pesos!$A$2:$A$20000," & cCod & "),0)"
        MontoUSD ws, r
    End If

    If esFondo Then
        If Not IsEmpty(f(0)) Then ws.Cells(r, CC_CY).Value = f(0)
        For j = 0 To N_VENT - 1
            If Not IsEmpty(f(1 + j)) Then
                cl = NumALetra(CC_W1 + j)
                ws.Cells(r, CC_W1 + j).Formula = _
                    "=IFERROR(INDEX(Fondos!$D$2:$I$500,MATCH(" & cCod & ",Fondos!$A$2:$A$500,0)," & (j + 1) & ")+" & _
                    Pase(cMon, cl) & ","""")"
            End If
        Next j
    Else
        ws.Cells(r, CC_CY).Formula = _
            "=IFERROR(AVERAGEIFS(Retornos!$R$2:$R$5000,Retornos!$A$2:$A$5000," & cCod & ",Retornos!$F$2:$F$5000,""YTD""),"""")"
        For j = 0 To N_VENT - 1
            cl = NumALetra(CC_W1 + j)
            ws.Cells(r, CC_W1 + j).Formula = _
                "=IFERROR(AVERAGEIFS(Retornos!$Q$2:$Q$5000,Retornos!$A$2:$A$5000," & cCod & ",Retornos!$F$2:$F$5000," & cl & "$" & CU_CAB & ")+" & _
                Pase(cMon, cl) & ","""")"
        Next j
    End If
    If EsNumero(cyMan) Then ws.Cells(r, CC_CY).Value = CDbl(cyMan)

    ws.Cells(r, CC_CYP).Formula = _
        "=IF(NOT(ISNUMBER(" & NumALetra(CC_CY) & r & ")),""""," & _
        NumALetra(CC_CY) & r & "+IF(" & cMon & "=""USD""," & SH_CFG & "!$C$15,0))"
    colYTD = NumALetra(CC_W1 + 4)
    ws.Cells(r, CC_ANN).Formula = _
        "=IF(NOT(ISNUMBER(" & colYTD & r & ")),""""," & _
        "IFERROR((1+" & colYTD & r & ")^(365/(" & colYTD & "$" & CU_FIN & "-" & _
        colYTD & "$" & CU_BASE & "))-1,""""))"
    ws.Cells(r, CC_MEN).Formula = _
        "=IF(NOT(ISNUMBER(" & NumALetra(CC_ANN) & r & ")),""""," & NumALetra(CC_ANN) & r & "/12)"
End Sub


Private Sub MontoUSD(ws As Worksheet, r As Long)
    ' USD mm = peso x total del fondo (BD_Pesos M1/M2) / tipo de cambio / 1e6.
    ' Misma formula en todas las filas: las grises suman pesos, asi que suman montos.
    Dim j As Long, cl As String, tot As String, tc As String
    tc = "IF(N(" & SH_CFG & "!$C$24)>0," & SH_CFG & "!$C$24," & SH_PES & "!$M$3)"
    For j = 0 To 1
        cl = NumALetra(CC_P1 + j) & r
        tot = SH_PES & "!$M$" & (j + 1)
        ws.Cells(r, CC_U1 + j).Formula = _
            "=IF(NOT(ISNUMBER(" & cl & ")),""""," & "IFERROR(" & cl & "*" & tot & "/" & tc & "/1000000,""""))"
    Next j
End Sub


Private Sub FilaCategoria(ws As Worksheet, r As Long, a As Long, b As Long)
    ' Fila de categoria: promedio ponderado por F1 de SU tramo a..b. El tramo
    ' se ve en la formula; no hay columnas escondidas.
    Dim j As Long, cW As String
    mDetalle = "CUADRO: categoria en la fila " & r
    If b < a Then Exit Sub
    cW = NumALetra(CC_P1)
    ws.Cells(r, CC_P1).Formula = "=SUM(" & cW & a & ":" & cW & b & ")"
    ws.Cells(r, CC_P1 + 1).Formula = "=SUM(" & NumALetra(CC_P1 + 1) & a & ":" & NumALetra(CC_P1 + 1) & b & ")"
    MontoUSD ws, r
    ws.Cells(r, CC_CY).Formula = PromTramo(a, b, CC_CY)
    ws.Cells(r, CC_CYP).Formula = PromTramo(a, b, CC_CYP)
    For j = 0 To N_VENT - 1
        ws.Cells(r, CC_W1 + j).Formula = PromTramo(a, b, CC_W1 + j)
    Next j
    ws.Cells(r, CC_MEN).Formula = PromTramo(a, b, CC_MEN)
    ws.Cells(r, CC_ANN).Formula = PromTramo(a, b, CC_ANN)
    ws.Cells(r, CC_DUR).Formula = PromTramo(a, b, CC_DUR)
    ws.Cells(r, CC_SPR).Formula = PromTramo(a, b, CC_SPR)
End Sub


Private Function PromTramo(a As Long, b As Long, col As Long) As String
    ' SUMPRODUCT(peso F1 x valor) / SUMPRODUCT(peso F1), solo donde hay valor.
    Dim w As String, x As String
    w = "$" & NumALetra(CC_P1) & "$" & a & ":$" & NumALetra(CC_P1) & "$" & b
    x = NumALetra(col) & "$" & a & ":" & NumALetra(col) & "$" & b
    PromTramo = "=IFERROR(SUMPRODUCT(" & w & "*ISNUMBER(" & x & ")," & x & ")/SUMPRODUCT(" & w & "*ISNUMBER(" & x & ")),"""")"
End Function


Private Sub FilaExplicita(ws As Worksheet, r As Long, filas As String, conPrima As Boolean)
    ' Fila que promedia OTRAS filas grises, nombradas una por una: el bloque
    ' sus categorias, el Total sus bloques, los SubTotales sus categorias.
    Dim j As Long, lst As Variant
    mDetalle = "CUADRO: fila agregada en " & r
    If Len(filas) = 0 Then Exit Sub
    lst = Split(filas, "|")
    ws.Cells(r, CC_P1).Formula = SumaFilas(lst, CC_P1)
    ws.Cells(r, CC_P1 + 1).Formula = SumaFilas(lst, CC_P1 + 1)
    MontoUSD ws, r
    ws.Cells(r, CC_CY).Formula = PromFilas(lst, CC_CY)
    ws.Cells(r, CC_CYP).Formula = PromFilas(lst, CC_CYP)
    For j = 0 To N_VENT - 1
        ws.Cells(r, CC_W1 + j).Formula = MasPrima(PromFilas(lst, CC_W1 + j), CC_W1 + j, conPrima)
    Next j
    ws.Cells(r, CC_MEN).Formula = MasPrima(PromFilas(lst, CC_MEN), CC_MEN, conPrima)
    ws.Cells(r, CC_ANN).Formula = MasPrima(PromFilas(lst, CC_ANN), CC_ANN, conPrima)
    ws.Cells(r, CC_DUR).Formula = PromFilas(lst, CC_DUR)
    ws.Cells(r, CC_SPR).Formula = PromFilas(lst, CC_SPR)
End Sub


Private Function SumaFilas(lst As Variant, col As Long) As String
    Dim i As Long, f As String
    For i = LBound(lst) To UBound(lst)
        f = f & IIf(Len(f) > 0, "+", "") & "N(" & NumALetra(col) & CStr(lst(i)) & ")"
    Next i
    SumaFilas = "=" & f
End Function


Private Function PromFilas(lst As Variant, col As Long) As String
    ' (F1_r1*X_r1 + F1_r2*X_r2 + ...) / (F1_r1 + F1_r2 + ...), contando solo
    ' las filas donde X es numero.
    Dim i As Long, arriba As String, abajo As String, w As String, x As String
    For i = LBound(lst) To UBound(lst)
        w = "$" & NumALetra(CC_P1) & CStr(lst(i))
        x = NumALetra(col) & CStr(lst(i))
        arriba = arriba & IIf(Len(arriba) > 0, "+", "") & "IF(ISNUMBER(" & x & ")," & w & "*" & x & ",0)"
        abajo = abajo & IIf(Len(abajo) > 0, "+", "") & "IF(ISNUMBER(" & x & ")," & w & ",0)"
    Next i
    PromFilas = "=IFERROR((" & arriba & ")/(" & abajo & "),"""")"
End Function


Private Function FilasManuales(ws As Worksheet, r As Long, man As Variant, manVar As Object, _
                               cod As String, nom As String, tipos As Object) As Long
    ' Filas manuales ancladas a este instrumento, debajo de el. Como quedan
    ' dentro del tramo de su categoria, SI entran al subtotal con el peso que
    ' tu escribas.
    Dim i As Long, fila As Variant, anc As String
    FilasManuales = r
    If Not IsArray(man) Then Exit Function
    For i = LBound(man) To UBound(man)
        fila = man(i)
        anc = CStr(fila(1))
        If NormTxt(anc) = NormTxt(nom) Or SinGuiones(anc) = SinGuiones(cod) Then
            FilasManuales = FilasManuales + 1
            EscribeManual ws, FilasManuales, fila, manVar
            tipos(FilasManuales) = "MAN"
        End If
    Next i
End Function


Private Sub EscribeManual(ws As Worksheet, r As Long, fila As Variant, manVar As Object)
    Dim j As Long, colYTD As String, vv As Variant
    mDetalle = "CUADRO: fila manual " & CStr(fila(0)) & " en " & r
    ws.Cells(r, CC_NOM).Value = CStr(fila(0))
    ws.Cells(r, CC_NOM).IndentLevel = 3
    ws.Cells(r, CC_MRE).Value = UCase$(Trim$(Txt(fila(13))))
    If EsNumero(fila(2)) Then ws.Cells(r, CC_P1).Value = CDbl(fila(2))
    If EsNumero(fila(3)) Then ws.Cells(r, CC_P1 + 1).Value = CDbl(fila(3))
    MontoUSD ws, r
    If EsNumero(fila(12)) Then
        ws.Cells(r, CC_CY).Value = CDbl(fila(12))
        If UCase$(Trim$(Txt(fila(13)))) = "USD" Then
            ws.Cells(r, CC_CYP).Formula = "=" & NumALetra(CC_CY) & r & "+" & SH_CFG & "!$C$15"
        Else
            ws.Cells(r, CC_CYP).Formula = "=" & NumALetra(CC_CY) & r
        End If
    End If
    For j = 0 To N_VENT - 1
        If EsNumero(fila(4 + j)) Then
            ws.Cells(r, CC_W1 + j).Value = CDbl(fila(4 + j))
        Else
            vv = VarVentana(manVar, CStr(fila(0)), j)
            If Not IsEmpty(vv) Then ws.Cells(r, CC_W1 + j).Value = CDbl(vv)
        End If
    Next j
    colYTD = NumALetra(CC_W1 + 4)
    If EsNumero(fila(11)) Then
        ws.Cells(r, CC_ANN).Value = CDbl(fila(11))
    Else
        ws.Cells(r, CC_ANN).Formula = _
            "=IF(NOT(ISNUMBER(" & colYTD & r & ")),""""," & _
            "IFERROR((1+" & colYTD & r & ")^(365/(" & colYTD & "$" & CU_FIN & "-" & _
            colYTD & "$" & CU_BASE & "))-1,""""))"
    End If
    If EsNumero(fila(10)) Then
        ws.Cells(r, CC_MEN).Value = CDbl(fila(10))
    Else
        ws.Cells(r, CC_MEN).Formula = _
            "=IF(NOT(ISNUMBER(" & NumALetra(CC_ANN) & r & ")),""""," & NumALetra(CC_ANN) & r & "/12)"
    End If
End Sub


'--------------------------------------------------------------------- formato
Private Sub FormatearCuadro(ws As Worksheet, ultFila As Long, tipos As Object)
    Dim r As Long, j As Long, tipo As String, rFila As Range

    If ultFila < CU_DAT Then Exit Sub
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 8

    ws.Columns(CC_NOM).ColumnWidth = 38
    ws.Columns(CC_COD).ColumnWidth = 14
    ws.Columns(CC_MRE).ColumnWidth = 5
    ws.Columns(CC_ISIN).ColumnWidth = 13
    ws.Columns(CC_RAT).ColumnWidth = 12
    ws.Range(ws.Columns(CC_DUR), ws.Columns(CC_SPR)).ColumnWidth = 9
    ws.Range(ws.Columns(CC_CY), ws.Columns(CC_CYP)).ColumnWidth = 9
    ws.Range(ws.Columns(CC_P1), ws.Columns(CC_U1 + 1)).ColumnWidth = 8
    ws.Columns(CC_SEP).ColumnWidth = 1.6
    ws.Range(ws.Columns(CC_W1), ws.Columns(CC_W1 + N_VENT - 1)).ColumnWidth = 9
    ws.Columns(CC_SEP2).ColumnWidth = 1.6
    ws.Range(ws.Columns(CC_MEN), ws.Columns(CC_ANN)).ColumnWidth = 10

    With ws.Range(ws.Cells(CU_FIN, CC_W1), ws.Cells(CU_BASE, CC_W1 + N_VENT - 1))
        .NumberFormat = "dd/mm/yy"
        .Font.Size = 7
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(150, 150, 150)
    End With
    ws.Range(ws.Cells(CU_DAT, CC_CY), ws.Cells(ultFila, CC_P1 + 1)).NumberFormat = "0.00%;-0.00%;"
    ws.Range(ws.Cells(CU_DAT, CC_W1), ws.Cells(ultFila, CC_W1 + N_VENT - 1)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(CU_DAT, CC_MEN), ws.Cells(ultFila, CC_ANN)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(CU_DAT, CC_DUR), ws.Cells(ultFila, CC_DUR)).NumberFormat = "0.00"
    ws.Range(ws.Cells(CU_DAT, CC_SPR), ws.Cells(ultFila, CC_SPR)).NumberFormat = "0"
    ws.Range(ws.Cells(CU_DAT, CC_U1), ws.Cells(ultFila, CC_U1 + 1)).NumberFormat = "#,##0.0"
    ws.Range(ws.Cells(CU_DAT, CC_MRE), ws.Cells(ultFila, CC_ANN)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(CU_DAT, CC_COD), ws.Cells(ultFila, CC_COD)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(CU_DAT, CC_COD), ws.Cells(ultFila, CC_ISIN)).Font.Size = 7

    ' Cabecera: rojo instrumento, negro CY / pesos / USD mm, rojo ventanas, gris Mensual / Ann
    With ws.Range(ws.Cells(CU_CAB, CC_NOM), ws.Cells(CU_CAB, CC_SPR))
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Cells(CU_CAB, CC_NOM).HorizontalAlignment = xlLeft
    With ws.Range(ws.Cells(CU_CAB, CC_CY), ws.Cells(CU_CAB, CC_U1 + 1))
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(CU_CAB, CC_W1), ws.Cells(CU_CAB, CC_W1 + N_VENT - 1))
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(CU_CAB, CC_MEN), ws.Cells(CU_CAB, CC_ANN))
        .Interior.Color = RGB(127, 127, 127)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Cells(CU_FIN, CC_W1).Interior.Color = RGB(255, 230, 100)
    ws.Cells(CU_FIN, CC_W1).Font.Bold = True
    ws.Cells(CU_FIN, CC_W1).Font.Color = RGB(0, 0, 0)

    With ws.Range(ws.Cells(CU_DAT, CC_NOM), ws.Cells(ultFila, CC_ANN)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With

    For r = CU_DAT To ultFila
        tipo = ""
        If tipos.Exists(r) Then tipo = CStr(tipos(r))
        Set rFila = ws.Range(ws.Cells(r, CC_NOM), ws.Cells(r, CC_ANN))
        Select Case tipo
            Case "BLQ"
                rFila.Interior.Color = RGB(128, 128, 128)
                rFila.Font.Color = RGB(255, 255, 255)
                rFila.Font.Bold = True
            Case "SUB"
                rFila.Interior.Color = RGB(217, 217, 217)
                rFila.Font.Bold = True
                ws.Cells(r, CC_NOM).IndentLevel = 1
            Case "TOT"
                rFila.Font.Bold = True
                rFila.Borders(xlEdgeTop).LineStyle = xlContinuous
                rFila.Borders(xlEdgeTop).Weight = xlMedium
                rFila.Borders(xlEdgeBottom).LineStyle = xlContinuous
                rFila.Borders(xlEdgeBottom).Weight = xlMedium
            Case "MAN"
                rFila.Font.Italic = True
            Case "REF"
                rFila.Font.Italic = True
                rFila.Font.Color = RGB(89, 89, 89)
            Case "ETQ"
                rFila.Borders.LineStyle = xlNone
                rFila.Font.Bold = True
                rFila.Borders(xlEdgeTop).LineStyle = xlContinuous
                rFila.Borders(xlEdgeTop).Weight = xlThin
            Case "NIV"
                rFila.Borders.LineStyle = xlNone
                rFila.Font.Bold = True
                rFila.Borders(xlEdgeTop).LineStyle = xlContinuous
                rFila.Borders(xlEdgeTop).Weight = xlThin
                rFila.Borders(xlEdgeBottom).LineStyle = xlContinuous
                rFila.Borders(xlEdgeBottom).Weight = xlThin
            Case "SEP", ""
                rFila.Borders.LineStyle = xlNone
        End Select
    Next r

    For j = 0 To N_VENT - 1
        Escala RangoIns(ws, CC_W1 + j, CU_DAT, ultFila, tipos)
    Next j
    Escala RangoIns(ws, CC_MEN, CU_DAT, ultFila, tipos)
    Escala RangoIns(ws, CC_ANN, CU_DAT, ultFila, tipos)
    CongelarPaneles ws, CU_CAB, CC_NOM
End Sub


Private Function RangoIns(ws As Worksheet, col As Long, r1 As Long, r2 As Long, tipos As Object) As Range
    Dim r As Long, u As Range, t As String
    For r = r1 To r2
        t = ""
        If tipos.Exists(r) Then t = CStr(tipos(r))
        If t = "INS" Or t = "OUT" Or t = "MAN" Or t = "REF" Then
            If u Is Nothing Then
                Set u = ws.Cells(r, col)
            Else
                Set u = Application.Union(u, ws.Cells(r, col))
            End If
        End If
    Next r
    Set RangoIns = u
End Function


Private Function NormMoneda(ByVal t As String) As String
    ' El FMS escribe "S/." y "US$". El cuadro agrupa por este texto, asi que
    ' se normaliza aqui y no en veinte sitios.
    Dim a As String
    a = NormTxt(t)
    a = Replace(Replace(Replace(a, ".", ""), "/", ""), " ", "")
    Select Case a
        Case "S", "SOLES", "PEN", "S1", "NUEVOSOL", "NUEVOSSOLES": NormMoneda = "PEN"
        Case "US$", "US", "USD", "USS", "DOLARES", "DOLAR", "$": NormMoneda = "USD"
        Case ""
            NormMoneda = "SIN MONEDA"
        Case Else
            NormMoneda = UCase$(Trim$(t))
    End Select
End Function





Private Function NormTxt(ByVal t As String) As String
    ' mayusculas, sin tildes, sin espacios de mas: para que "Ilíquidos" e
    ' "ILIQUIDOS" sean lo mismo y no dependa de como lo escriba el FMS
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





Private Function TotalAnexo(va As Variant, aMap As Object, fondo As String, codTot As String) As Double
    Dim i As Long
    If Not IsArray(va) Then Exit Function
    If aMap Is Nothing Then Exit Function
    ' OJO: el texto del anexo trae un espacio delante. Sin el Trim$, el InStr
    ' devuelve 2 en vez de 1 y ninguna fila cuadra. La radiografia del Mapa si
    ' lo hacia, por eso encontraba los totales y esta funcion no.
    For i = 1 To UBound(va, 1)
        If MismoFondo(Txt(va(i, aMap("fondo"))), fondo) Then
            If InStr(1, Trim$(Txt(va(i, aMap("codigo")))), codTot, vbTextCompare) = 1 Then
                TotalAnexo = Num(va(i, aMap("monto")))
                Exit Function
            End If
        End If
    Next i
End Function





Private Sub EscribirPesos(datos As Object, totF1 As Double, totF2 As Double, dFms As Double)
    Dim ws As Worksheet, out() As Variant, k As Variant, n As Long, arr As Variant
    Dim cab As Variant

    Set ws = HojaOCrea(SH_PES)
    ws.Cells.Clear

    cab = Array("Codigo SBS", "Nemonico", "Emisor", "Moneda", "Val_total F1", _
                "Val_total F2", "Peso F1", "Peso F2", "Fecha FMS", "Vencimiento (FMS)")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PE_COLS)).Value = cab
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PE_COLS)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PE_COLS)).Interior.Color = RGB(217, 217, 217)

    If datos.Count = 0 Then Exit Sub
    ReDim out(1 To datos.Count, 1 To PE_COLS)
    n = 0
    For Each k In datos.Keys
        n = n + 1
        arr = datos(k)
        out(n, 1) = "'" & CStr(k)
        out(n, 2) = arr(0)
        out(n, 3) = arr(1)
        out(n, 4) = NormMoneda(CStr(arr(2)))
        out(n, 5) = Num(arr(3))
        out(n, 6) = Num(arr(4))
        out(n, 7) = IIf(totF1 > 0, Num(arr(3)) / totF1, 0)
        out(n, 8) = IIf(totF2 > 0, Num(arr(4)) / totF2, 0)
        out(n, 9) = dFms
        If Not IsEmpty(arr(5)) Then out(n, 10) = arr(5)
    Next k

    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, PE_COLS)).Value = out
    ws.Range(ws.Cells(2, 5), ws.Cells(n + 1, 6)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(2, 7), ws.Cells(n + 1, 8)).NumberFormat = "0.000%"
    ws.Range(ws.Cells(2, 9), ws.Cells(n + 1, 10)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, PE_COLS)).EntireColumn.AutoFit
    CongelarPaneles ws, 1, 1


    ' El total de cada fondo (Anexo I) se guarda a la vista: de el sale el
    ' monto en USD mm del cuadro, y asi se puede auditar sin abrir el FMS.
    ws.Range("L1").Value = "Total Fondo 1 (PEN)"
    ws.Range("L2").Value = "Total Fondo 2 (PEN)"
    ws.Range("L1:L2").Font.Bold = True
    ws.Range("M1").Value = totF1
    ws.Range("M2").Value = totF2
    ws.Range("M1:M2").NumberFormat = "#,##0"
    TotalesFondo
End Sub





Private Sub GuardarTC(tc As Double)
    ' Deja el tipo de cambio del FMS en BD_Pesos M3, a la vista, al lado de los
    ' totales de cada fondo. Si el FMS no lo trajo, no pisa lo que ya habia.
    Dim ws As Worksheet
    If Not HojaExiste(SH_PES) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_PES)
    ws.Range("L3").Value = "Tipo de cambio USD/PEN (FMS)"
    ws.Range("L3").Font.Bold = True
    If tc > 0 Then ws.Range("M3").Value = tc
    ws.Range("M3").NumberFormat = "0.0000"
End Sub





Private Sub TotalesFondo()
    ' Deja en BD_Pesos M1 y M2 el total de cada fondo y define peTotF1/peTotF2.
    ' Si CargarFMS no los escribio (libro viejo), los deduce: para cualquier
    ' instrumento, total = Val_total / Peso. No necesita reabrir el FMS.
    Dim ws As Worksheet, i As Long, ult As Long
    Dim t1 As Double, t2 As Double, pe As Double

    If Not HojaExiste(SH_PES) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_PES)
    t1 = Num(ws.Range("M1").Value)
    t2 = Num(ws.Range("M2").Value)

    If t1 <= 0 Or t2 <= 0 Then
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To ult
            If t1 <= 0 Then
                pe = Num(ws.Cells(i, 7).Value)
                If pe > 0 Then t1 = Num(ws.Cells(i, 5).Value) / pe
            End If
            If t2 <= 0 Then
                pe = Num(ws.Cells(i, 8).Value)
                If pe > 0 Then t2 = Num(ws.Cells(i, 6).Value) / pe
            End If
            If t1 > 0 And t2 > 0 Then Exit For
        Next i
        ws.Range("L1").Value = "Total Fondo 1 (PEN)"
        ws.Range("L2").Value = "Total Fondo 2 (PEN)"
        ws.Range("L1:L2").Font.Bold = True
        If t1 > 0 Then ws.Range("M1").Value = t1
        If t2 > 0 Then ws.Range("M2").Value = t2
        ws.Range("M1:M2").NumberFormat = "#,##0"
    End If

End Sub





Public Sub CargarFMSForzado()
    ' Relee el FMS aunque BD_Pesos ya lo tenga. Usalo si corrigieron el archivo.
    mForzar = True
    CargarFMS
    mForzar = False
End Sub





Private Function FmsYaCargado(dFms As Double) As Boolean
    Dim ws As Worksheet, ult As Long
    If Not HojaExiste(SH_PES) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_PES)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function
    If Not EsNumero(ws.Cells(2, 9).Value) Then Exit Function
    FmsYaCargado = (CLng(ws.Cells(2, 9).Value) = CLng(dFms))
End Function





Public Sub NormalizarExcel()
    ' Si una macro se corta a medias (Esc + Finalizar, o un error sin manejar),
    ' Excel se queda en calculo MANUAL, sin refrescar pantalla y sin eventos.
    ' Y eso es un ajuste de Excel entero, no del libro: todos los archivos que
    ' abras despues quedan igual. Esto lo devuelve a la normalidad.
    Dim antes As String

    antes = "Calculo: " & IIf(Application.Calculation = xlCalculationAutomatic, _
                              "automatico", "MANUAL")

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Cursor = xlDefault

    MsgBox "Excel normalizado." & vbCrLf & vbCrLf & _
           "Como estaba:  " & antes & vbCrLf & vbCrLf & _
           "Ahora: calculo automatico, pantalla y eventos activos." & vbCrLf & vbCrLf & _
           "Corre esto cada vez que cortes una macro con Esc.", _
           vbInformation, "NormalizarExcel"
End Sub





Private Sub BotonesEnColumna(ws As Worksheet, lista As Variant, izq As Double, _
                             arriba As Double, ancho As Double, alto As Double, _
                             destacar As Boolean)
    ' lista viene de a pares: nombre de macro, texto del boton.
    Dim i As Long, k As Long, sh As Shape, y As Double
    y = arriba
    k = 0
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
        k = k + 1
    Next i
End Sub





Private Function BuscarFms(carpeta As String, dCorte As Double, ByRef ruta As String) As Double
    Dim i As Long, p As String, ext As Variant
    For i = 0 To 20
        For Each ext In Array(".xlsx", ".xls", ".xlsb")
            p = carpeta & Application.PathSeparator & "FMS_" & _
                Format$(CDate(dCorte - i), "yyyymmdd") & CStr(ext)
            If Dir(p) <> "" Then
                ruta = p
                BuscarFms = dCorte - i
                Exit Function
            End If
        Next ext
    Next i
End Function





Private Function TipoCambioFms(wb As Workbook, wsCfg As Worksheet) As Double
    ' El FMS viene en SOLES. En su hoja T.Cambio esta cuantos soles vale cada
    ' moneda: la fila del simbolo US$ trae el dolar. De ahi sale el USD mm.
    ' Hoja en Config C41, simbolo en C45.
    Dim ws As Worksheet, hoja As String, simb As String
    Dim v As Variant, ult As Long, i As Long, j As Long
    Dim colSim As Long, colVal As Long, cab As String, x As Double

    hoja = Txt(wsCfg.Range("C41").Value): If Len(hoja) = 0 Then hoja = "T.Cambio"
    simb = Txt(wsCfg.Range("C42").Value): If Len(simb) = 0 Then simb = "US$"

    On Error Resume Next
    Set ws = wb.Worksheets(hoja)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function
    If ult > 500 Then ult = 500
    v = ws.Range(ws.Cells(1, 1), ws.Cells(ult, 8)).Value

    ' La cabecera manda: sim_moneda y com_vrtc. Si no la encuentro, C y D.
    For j = 1 To 8
        cab = LCase$(Txt(v(1, j)))
        If InStr(cab, "sim") > 0 Then
            If colSim = 0 Then colSim = j
        End If
        If InStr(cab, "vrtc") > 0 Then
            If colVal = 0 Then colVal = j
        End If
    Next j
    If colSim = 0 Then colSim = 3
    If colVal = 0 Then colVal = 4

    For i = 2 To ult
        If NormTxt(Txt(v(i, colSim))) = NormTxt(simb) Then
            x = Num(v(i, colVal))
            If x > 0 Then TipoCambioFms = x
            Exit Function
        End If
    Next i
End Function





Public Sub ProbarTipoCambio()
    ' Abre el FMS en SOLO LECTURA y dice que encontro en la hoja del tipo de
    ' cambio: que hojas hay, que cabeceras, y las primeras monedas con su
    ' valor. Es para cuando el USD mm sale en blanco y no se sabe por que.
    Dim wsCfg As Worksheet, wb As Workbook, ws As Worksheet
    Dim carpeta As String, ruta As String, hoja As String, simb As String
    Dim dCorte As Double, dFms As Double, tc As Double
    Dim msg As String, i As Long, j As Long, ult As Long, nh As Long
    Dim alertPrev As Boolean

    On Error GoTo Falla
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    carpeta = SinBarra(Txt(wsCfg.Range("C4").Value))
    hoja = Txt(wsCfg.Range("C41").Value): If Len(hoja) = 0 Then hoja = "T.Cambio"
    simb = Txt(wsCfg.Range("C42").Value): If Len(simb) = 0 Then simb = "US$"
    dCorte = CorteEfectivo(wsCfg)
    dFms = BuscarFms(carpeta, dCorte, ruta)
    If dFms = 0 Then
        Aviso "No encontre ningun FMS <= " & Format$(CDate(dCorte), "dd/mm/yyyy") & _
              vbCrLf & "Carpeta (Config C4): " & carpeta, "ProbarTipoCambio"
        Exit Sub
    End If

    alertPrev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    msg = "FMS: " & ruta & vbCrLf & vbCrLf
    msg = msg & "Busco la hoja [" & hoja & "] (Config C41)" & vbCrLf
    On Error Resume Next
    Set ws = wb.Worksheets(hoja)
    On Error GoTo Falla

    If ws Is Nothing Then
        msg = msg & "NO EXISTE esa hoja. Las que si estan:" & vbCrLf
        For nh = 1 To wb.Worksheets.Count
            If nh <= 30 Then msg = msg & "   " & wb.Worksheets(nh).Name & vbCrLf
        Next nh
        msg = msg & vbCrLf & "Copia el nombre exacto a Config C41."
    Else
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        msg = msg & "La encontre. Filas con dato: " & ult & vbCrLf & vbCrLf
        msg = msg & "Cabecera (fila 1):" & vbCrLf
        For j = 1 To 6
            msg = msg & "   col " & NumALetra(j) & ": [" & Txt(ws.Cells(1, j).Value) & "]" & vbCrLf
        Next j
        msg = msg & vbCrLf & "Primeras filas:" & vbCrLf
        For i = 2 To 6
            If i <= ult Then
                msg = msg & "   [" & Txt(ws.Cells(i, 3).Value) & "] = " & _
                      Txt(ws.Cells(i, 4).Value) & vbCrLf
            End If
        Next i
        tc = TipoCambioFms(wb, wsCfg)
        msg = msg & vbCrLf & "Busco el simbolo [" & simb & "] (Config C42)" & vbCrLf
        If tc > 0 Then
            msg = msg & "RESULTADO: " & Format$(tc, "0.0000") & "   <- esto va a BD_Pesos M3"
        Else
            msg = msg & "NO LO ENCONTRE. Mira arriba como esta escrito el dolar " & _
                  "y copialo igualito a Config C42."
        End If
    End If

    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.DisplayAlerts = alertPrev
    If tc > 0 Then GuardarTC tc
    Aviso msg, "ProbarTipoCambio"
    Exit Sub

Falla:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Aviso "ProbarTipoCambio se detuvo." & vbCrLf & vbCrLf & Err.Description, "ProbarTipoCambio"
End Sub





Private Function MapaFms(wsCfg As Worksheet) As Object
    Dim d As Object
    Set d = NuevoDic()
    d("fondo") = Letra(Txt(wsCfg.Range("C27").Value), 1)
    d("asset") = Letra(Txt(wsCfg.Range("C28").Value), 3)
    d("codsbs") = Letra(Txt(wsCfg.Range("C29").Value), 4)
    d("emisor") = Letra(Txt(wsCfg.Range("C30").Value), 5)
    d("nemonico") = Letra(Txt(wsCfg.Range("C31").Value), 7)
    d("moneda") = Letra(Txt(wsCfg.Range("C32").Value), 8)
    d("valtotal") = Letra(Txt(wsCfg.Range("C33").Value), 10)
    d("vencimiento") = Letra(Txt(wsCfg.Range("C34").Value), 12)
    Set MapaFms = d
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




Private Sub Aviso(texto As String, titulo As String)
    ' En ActualizarTodo junta todo en un solo dialogo al final.
    If mSilencio Then
        If Len(mResumen) > 0 Then mResumen = mResumen & vbCrLf & String$(46, "-") & vbCrLf
        mResumen = mResumen & "[" & titulo & "]" & vbCrLf & texto
    Else
        MsgBox texto, vbInformation, titulo
    End If
End Sub





Private Sub Paso(texto As String)
    mPasos = mPasos & "  - " & texto & " (" & Format$(Timer - mT0, "0.0") & " s)" & vbCrLf
    mDetalle = texto
    Application.StatusBar = "Iliquidos: " & texto & "..."
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
    mFallo = True
    mSilencio = False
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
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





Public Sub GuardarLog(texto As String)
    Dim ws As Worksheet, partes As Variant, i As Long, f As Integer, ruta As String
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
    If Len(ThisWorkbook.Path) > 0 Then
        ruta = ThisWorkbook.Path & Application.PathSeparator & "Log_Iliquidos.txt"
        f = FreeFile
        Open ruta For Output As #f
        Print #f, "=== " & Format$(Now, "dd/mm/yyyy hh:mm:ss") & " ==="
        Print #f, texto
        Close #f
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
    If ws.Visible <> xlSheetVisible Then Exit Sub
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





Private Function MismoFondo(ByVal a As String, ByVal b As String) As Boolean
    ' "01" = "1" = 1 = " 01 " = "F1" ... todo lo que en el fondo es el mismo fondo.
    Dim x As String, y As String
    x = SoloDigitos(a)
    y = SoloDigitos(b)
    If Len(x) > 0 And Len(y) > 0 Then
        MismoFondo = (CDbl(x) = CDbl(y))
    Else
        MismoFondo = (NormFondo(a) = NormFondo(b))
    End If
End Function





Private Function SoloDigitos(ByVal t As String) As String
    Dim i As Long, c As String, out As String
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c >= "0" And c <= "9" Then out = out & c
    Next i
    If Len(out) > 15 Then out = Left$(out, 15)
    SoloDigitos = out
End Function





Private Function NormFondo(ByVal s As String) As String
    ' "09" -> "9",  "01" -> "1"
    s = Trim$(s)
    Do While Len(s) > 1 And Left$(s, 1) = "0"
        s = Mid$(s, 2)
    Loop
    NormFondo = s
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
        ' OJO: IsNumeric() de una fecha devuelve False. Por eso el vencimiento
        ' salia 0 y se veia como 00/01/1900. Una fecha SI es un numero.
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





Private Function Num(v As Variant) As Double
    If EsNumero(v) Then Num = CDbl(v)
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
    ' el FMS trae AAAAMMDD como texto; el vector trae fecha de Excel
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





Private Function SumaDic(datos As Object, cual As Long) As Double
    Dim k As Variant, arr As Variant
    For Each k In datos.Keys
        arr = datos(k)
        SumaDic = SumaDic + Num(arr(2 + cual))
    Next k
End Function




Private Function DiagMontos() As String
    ' Las tres piezas de la columna USD mm. Si alguna falta el cuadro sale en
    ' blanco, y antes no habia forma de saber cual era sin abrir las formulas.
    Dim ws As Worksheet, t1 As Double, t2 As Double, tc As Double, man As Double
    Dim msg As String

    If Not HojaExiste(SH_PES) Then
        DiagMontos = "USD mm: no existe BD_Pesos. Corre CargarFMS." & vbCrLf
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(SH_PES)
    t1 = Num(ws.Range("M1").Value)
    t2 = Num(ws.Range("M2").Value)
    tc = Num(ws.Range("M3").Value)
    man = Num(ThisWorkbook.Worksheets(SH_CFG).Range("C24").Value)

    msg = "USD mm:  Total F1 " & IIf(t1 > 0, Format$(t1 / 1000000, "#,##0") & " mn", "FALTA") & _
        "   Total F2 " & IIf(t2 > 0, Format$(t2 / 1000000, "#,##0") & " mn", "FALTA") & vbCrLf
    If man > 0 Then
        msg = msg & "         TC " & Format$(man, "0.0000") & " (a mano, Config C24)" & vbCrLf
    ElseIf tc > 0 Then
        msg = msg & "         TC " & Format$(tc, "0.0000") & " (del FMS)" & vbCrLf
    Else
        msg = msg & "         TC FALTA: corre CargarFMSForzado, o escribelo en Config C24." & vbCrLf
    End If
    DiagMontos = msg
End Function





Private Function EsRef(arr As Variant) As Boolean
    ' Novena posicion del array de LeerInstrumentos. Se pregunta con UBound
    ' para no reventar con arrays viejos que no la traen.
    If Not IsArray(arr) Then Exit Function
    If UBound(arr) < 8 Then Exit Function
    If VarType(arr(8)) = vbBoolean Then EsRef = CBool(arr(8))
End Function





Private Function MonedaBloque(cat As String, mon As String) As String
    ' Hay categorias que van siempre al bloque PEN, aunque el instrumento este
    ' en dolares: en el cuadro se muestran ya pasadas a soles. La moneda real
    ' se guarda aparte, porque es la que decide si se le suma el spread.
    Dim cfgPen As String, p As Variant
    MonedaBloque = mon
    On Error Resume Next
    cfgPen = Txt(ThisWorkbook.Worksheets(SH_CFG).Range("C16").Value)
    On Error GoTo 0
    If Len(Trim$(cfgPen)) = 0 Then Exit Function
    For Each p In Split(cfgPen, "|")
        If Len(Trim$(CStr(p))) > 0 Then
            If NormTxt(cat) = NormTxt(CStr(p)) Then
                MonedaBloque = "PEN"
                Exit Function
            End If
        End If
    Next p
End Function





Private Sub ArmarRetornos(ins As Object, orden As Variant, dCorte As Double, wsCfg As Worksheet)
    Dim ws As Worksheet, r As Long, i As Long, j As Long
    Dim vent As Variant, bases() As Double
    Dim arr As Variant, cab As Variant
    Dim tol As Long

    vent = Array("WTD", "5D", "MTD", "MayoTD", "YTD", "FY")
    bases = BasesVentana(wsCfg, dCorte)
    tol = CLng(NumDef(wsCfg.Range("C9").Value, 5))

    Set ws = HojaOCrea(SH_RET)
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0
    ws.Cells.Clear

    cab = Array("Codigo SBS", "Nemonico", "Nombre", "Categoria", "Moneda", "Ventana", _
                "Fecha base", "Fecha fin", "Base efectiva", "Fin efectiva", _
                "P base", "P fin", "Dias", "Tasa cupon", _
                "Ret. price", "Ret. cupon", "SUMA", "Curr. yield", "Flag")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 19)).Value = cab
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 19)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 19)).Interior.Color = RGB(64, 64, 64)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 19)).Font.Color = RGB(255, 255, 255)

    r = 1
    If IsArray(orden) Then
        For i = LBound(orden) To UBound(orden)
            arr = ins(orden(i))
            For j = 0 To N_VENT - 1
                r = r + 1
                mDetalle = "Retornos: " & arr(1) & " / " & vent(j)
                ws.Cells(r, 1).Value = "'" & arr(0)
                ws.Cells(r, 2).Formula = "=IFERROR(INDEX(BD_Pesos!$B$2:$B$20000,MATCH($A" & r & ",BD_Pesos!$A$2:$A$20000,0)),"""")"
                ws.Cells(r, 3).Value = arr(1)
                ws.Cells(r, 4).Value = arr(2)
                ws.Cells(r, 5).Value = arr(3)
                ws.Cells(r, 6).Value = vent(j)
                ws.Cells(r, 7).Value = bases(j)
                ws.Cells(r, 8).Value = dCorte
                ' Si no hay precio en la fecha base, se arranca desde el primer
                ' precio que exista dentro de la ventana. La columna Flag lo avisa.
                ws.Cells(r, 9).Formula = _
                    "=IFERROR(IF(MAXIFS(Precios!$A$2:$A$60000,Precios!$B$2:$B$60000,$A" & r & ",Precios!$A$2:$A$60000,""<=""&G" & r & ")>0," & _
                    "MAXIFS(Precios!$A$2:$A$60000,Precios!$B$2:$B$60000,$A" & r & ",Precios!$A$2:$A$60000,""<=""&G" & r & ")," & _
                    "IF(MINIFS(Precios!$A$2:$A$60000,Precios!$B$2:$B$60000,$A" & r & ",Precios!$A$2:$A$60000,""<=""&H" & r & ")>0," & _
                    "MINIFS(Precios!$A$2:$A$60000,Precios!$B$2:$B$60000,$A" & r & ",Precios!$A$2:$A$60000,""<=""&H" & r & "),"""")),"""")"
                ws.Cells(r, 10).Formula = _
                    "=IFERROR(MAXIFS(Precios!$A$2:$A$60000,Precios!$B$2:$B$60000,$A" & r & ",Precios!$A$2:$A$60000,""<=""&H" & r & "),"""")"
                ws.Cells(r, 11).Formula = _
                    "=IFERROR(AVERAGEIFS(Precios!$E$2:$E$60000,Precios!$B$2:$B$60000,$A" & r & ",Precios!$A$2:$A$60000,I" & r & "),"""")"
                ws.Cells(r, 12).Formula = _
                    "=IFERROR(AVERAGEIFS(Precios!$E$2:$E$60000,Precios!$B$2:$B$60000,$A" & r & ",Precios!$A$2:$A$60000,J" & r & "),"""")"
                ws.Cells(r, 13).Formula = "=IF(OR(I" & r & "="""",J" & r & "=""""),"""",J" & r & "-I" & r & ")"
                ws.Cells(r, 14).Formula = _
                    "=IFERROR(AVERAGEIFS(Precios!$G$2:$G$60000,Precios!$B$2:$B$60000,$A" & r & ",Precios!$A$2:$A$60000,J" & r & "),"""")"
                ws.Cells(r, 15).Formula = "=IFERROR((L" & r & "-K" & r & ")/K" & r & ","""")"
                ws.Cells(r, 16).Formula = "=IFERROR(N" & r & "*M" & r & "/365/K" & r & ","""")"
                ws.Cells(r, 17).Formula = "=IFERROR(O" & r & "+P" & r & ","""")"
                ws.Cells(r, 18).Formula = "=IFERROR(N" & r & "/L" & r & ","""")"
                ws.Cells(r, 19).Formula = _
                    "=IF(OR(K" & r & "="""",L" & r & "=""""),""REVISAR: falta precio""," & _
                    "IF(I" & r & ">G" & r & ",""REVISAR: sin precio al ""&TEXT(G" & r & _
                    ",""dd/mm/yy"")&""; la ventana arranca el ""&TEXT(I" & r & ",""dd/mm/yy"")&" & _
                    """ (""&M" & r & "&"" dias en vez de ""&(H" & r & "-G" & r & ")&"")""," & _
                    "IF(G" & r & "-I" & r & ">" & tol & ",""REVISAR: la base usa un vector ""&(G" & r & _
                    "-I" & r & ")&"" dias anterior""," & _
                    "IF(H" & r & "-J" & r & ">" & tol & ",""REVISAR: el fin usa un vector ""&(H" & r & _
                    "-J" & r & ")&"" dias anterior"",""""))))"
            Next j
        Next i
    End If

    ws.Range(ws.Cells(2, 7), ws.Cells(r, 10)).NumberFormat = "dd/mm/yy"
    ws.Range(ws.Cells(2, 11), ws.Cells(r, 12)).NumberFormat = "0.0000"
    ws.Range(ws.Cells(2, 14), ws.Cells(r, 14)).NumberFormat = "0.000"
    ws.Range(ws.Cells(2, 15), ws.Cells(r, 18)).NumberFormat = "0.00%"
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 14
    ws.Columns(3).ColumnWidth = 32
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 9
    ws.Columns(6).ColumnWidth = 9
    ws.Range(ws.Columns(7), ws.Columns(18)).ColumnWidth = 11
    ws.Columns(19).ColumnWidth = 46
    CongelarPaneles ws, 1, 3

End Sub





Private Function RestarHabiles(ByVal dFin As Double, ByVal cuantos As Long) As Double
    ' Resta dias habiles (lunes a viernes). Los feriados no se conocen, pero
    ' como el precio se busca hacia atras desde la base, un feriado no rompe:
    ' simplemente se toma el ultimo vector que exista antes.
    Dim i As Long
    RestarHabiles = dFin
    If cuantos <= 0 Then Exit Function
    For i = 1 To cuantos
        Do
            RestarHabiles = RestarHabiles - 1
        Loop While Weekday(CDate(RestarHabiles), vbMonday) > 5
    Next i
End Function





Private Function BasesVentana(wsCfg As Worksheet, dCorte As Double) As Double()
    Dim b(0 To 5) As Double, dc As Date
    dc = CDate(dCorte)
    ' WTD: el domingo anterior. Como el precio se busca hacia atras, el que
    ' termina usando es el del viernes de la semana pasada.
    b(0) = dCorte - (Weekday(dc, vbMonday) - 1) - 1
    ' 5D: cinco dias HABILES hacia atras, no calendario.
    b(1) = RestarHabiles(dCorte, CLng(NumDef(wsCfg.Range("C8").Value, 5)))
    b(2) = CDbl(DateSerial(Year(dc), Month(dc), 0))              ' MTD
    If IsDate(wsCfg.Range("C6").Value) Then
        b(3) = CDbl(CDate(wsCfg.Range("C6").Value)) - 1          ' MayoTD
    Else
        b(3) = dCorte
    End If
    b(4) = CDbl(DateSerial(Year(dc), 1, 1)) - 1                  ' YTD
    If IsDate(wsCfg.Range("C7").Value) Then
        b(5) = CDbl(CDate(wsCfg.Range("C7").Value))              ' FY
    Else
        b(5) = CDbl(DateSerial(Year(dc) - 1, 10, 31))
    End If
    BasesVentana = b
End Function




Public Sub CrearManual()
    ' Hoja para las filas que escribes a mano y que deben sobrevivir a
    ' ARMAR CUADRO (por ejemplo "Deals exPeru" debajo de Moneda Patria).
    ' Si la hoja ya existe solo repone la cabecera: tus filas no se tocan.
    Dim ws As Worksheet, cab As Variant

    On Error GoTo Falla
    Set ws = HojaOCrea(SH_MAN)

    cab = Array("Nombre de la fila", "Va debajo de", "F1", "F2", _
                "WTD", "5D", "MTD", "MayoTD", "YTD", "FY", _
                "Mensual YTD", "Ann YTD", "Curr. Yield", "Moneda")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 14)).Value = cab
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 14))
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns(1).ColumnWidth = 32
    ws.Columns(2).ColumnWidth = 32
    ws.Range(ws.Columns(3), ws.Columns(13)).ColumnWidth = 11
    ws.Columns(14).ColumnWidth = 9
    ws.Range(ws.Cells(2, 3), ws.Cells(500, 13)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(2, 14), ws.Cells(500, 14)).NumberFormat = "@"

    ws.Range("N2:N500").ClearContents
    ws.Range("P1").Value = "COMO SE USA"
    ws.Range("P1").Font.Bold = True
    ws.Range("P2").Value = "A  Nombre: el texto que quieres ver en el cuadro."
    ws.Range("P3").Value = "B  Va debajo de: el nombre del instrumento tal como sale en el " & _
                           "cuadro, o su codigo SBS."
    ws.Range("P4").Value = "C a M: los numeros. Escribe 0.0457 para 4.57%."
    ws.Range("P5").Value = "N  Moneda: PEN o USD. Solo decide el Curr. Yield en soles " & _
                           "(le suma el spread de Config C15). Vacio = PEN."
    ws.Range("P6").Value = "Lo que dejes vacio de E a J se calcula desde la hoja " & _
                           SH_MVA & ", si ahi hay marcas de esta fila."
    ws.Range("P7").Value = "Si dejas vacios Mensual YTD y Ann YTD, se calculan solos desde el YTD."
    ws.Range("P8").Value = "A los retornos NO se les suma el spread: sale lo que escribas."
    ws.Range("P9").Value = "Estas filas NO entran en ningun promedio, subtotal ni total."
    ws.Range("P10").Value = "Una fila por linea. Puedes poner varias debajo del mismo instrumento."
    ws.Range("P1:P10").Font.Color = RGB(100, 100, 100)
    ws.Columns(16).ColumnWidth = 84

    ' --- hoja de las variaciones, para no tener que calcular las ventanas
    Set ws = HojaOCrea(SH_MVA)
    ws.Range("A1").Value = "Nombre de la fila"
    ws.Range("B1").Value = "Fecha"
    ws.Range("C1").Value = "Var Ajustada"
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns(1).ColumnWidth = 32
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 14
    ws.Range("B2:B5000").NumberFormat = "dd/mm/yyyy"
    ws.Range("C2:C5000").NumberFormat = "0.0000%"

    ws.Range("E1").Value = "COMO SE USA"
    ws.Range("E1").Font.Bold = True
    ws.Range("E2").Value = "Una linea por marca: el mismo nombre que pusiste en la " & _
                           "columna A de la hoja Manual, la fecha y la variacion."
    ws.Range("E3").Value = "Escribe 0.0125 para 1.25%. Pega todas las que tengas, " & _
                           "en cualquier orden."
    ws.Range("E4").Value = "El macro encadena cada ventana solo: " & _
                           "EXP(SUMA(LN(1+var)))-1, igual que con los fondos de Marcas."
    ws.Range("E5").Value = "Manda lo que escribas a mano en la hoja Manual. " & _
                           "Lo que dejes vacio ahi se llena desde aqui."
    ws.Range("E6").Value = "Una ventana sin ninguna marca queda vacia, no en cero."
    ws.Range("E1:E6").Font.Color = RGB(100, 100, 100)
    ws.Columns(5).ColumnWidth = 84

    MsgBox "Hojas '" & SH_MAN & "' y '" & SH_MVA & "' listas." & vbCrLf & vbCrLf & _
           "En " & SH_MAN & " pones la fila y donde va." & vbCrLf & _
           "En " & SH_MVA & " pegas las variaciones ajustadas." & vbCrLf & vbCrLf & _
           "Despues corre ARMAR CUADRO.", vbInformation, "CrearManual"
    Exit Sub

Falla:
    MsgBox "CrearManual se detuvo." & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "CrearManual"
End Sub





Private Function LeerManual() As Variant
    ' Devuelve Empty si no hay hoja o no hay filas.
    ' Cada elemento: Array(nombre, ancla, F1, F2, w0..w5, mensual, ann, cy, moneda)
    Dim ws As Worksheet, ult As Long, i As Long, nn As Long
    Dim res() As Variant, nom As String, anc As String

    If Not HojaExiste(SH_MAN) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_MAN)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function

    ReDim res(0 To ult - 2)
    nn = 0
    For i = 2 To ult
        nom = Txt(ws.Cells(i, 1).Value)
        anc = Txt(ws.Cells(i, 2).Value)
        If Len(nom) > 0 And Len(anc) > 0 Then
            res(nn) = Array(nom, anc, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, _
                            ws.Cells(i, 5).Value, ws.Cells(i, 6).Value, _
                            ws.Cells(i, 7).Value, ws.Cells(i, 8).Value, _
                            ws.Cells(i, 9).Value, ws.Cells(i, 10).Value, _
                            ws.Cells(i, 11).Value, ws.Cells(i, 12).Value, _
                            ws.Cells(i, 13).Value, ws.Cells(i, 14).Value)
            nn = nn + 1
        End If
    Next i
    If nn = 0 Then Exit Function
    ReDim Preserve res(0 To nn - 1)
    LeerManual = res
End Function





Private Function LeerManualVar(dCorte As Double) As Object
    ' Hoja Manual_Var: Nombre | Fecha | Var Ajustada, una linea por marca.
    ' Encadena igual que ArmarFondos: EXP(SUMA(LN(1+var)))-1 por ventana.
    ' Devuelve un diccionario nombre -> Array(ln0..ln5, n0..n5).
    Dim ws As Worksheet, ult As Long, i As Long, j As Long
    Dim d As Object, k As String, arr As Variant
    Dim bases() As Double, dFe As Double, v As Double

    Set d = NuevoDic()
    Set LeerManualVar = d
    If Not HojaExiste(SH_MVA) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_MVA)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function

    bases = BasesVentana(ThisWorkbook.Worksheets(SH_CFG), dCorte)

    For i = 2 To ult
        k = NormTxt(Txt(ws.Cells(i, 1).Value))
        If Len(k) > 0 Then
            If EsNumero(ws.Cells(i, 2).Value) Then
                If EsNumero(ws.Cells(i, 3).Value) Then
                    dFe = CDbl(ws.Cells(i, 2).Value)
                    v = CDbl(ws.Cells(i, 3).Value)
                    If v > -1 Then
                        If Not d.Exists(k) Then
                            d(k) = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                        End If
                        arr = d(k)
                        For j = 0 To N_VENT - 1
                            If dFe > bases(j) Then
                                If dFe <= dCorte Then
                                    arr(j) = CDbl(arr(j)) + Log(1 + v)
                                    arr(6 + j) = CLng(arr(6 + j)) + 1
                                End If
                            End If
                        Next j
                        d(k) = arr
                    End If
                End If
            End If
        End If
    Next i
End Function





Private Function VarVentana(manVar As Object, nom As String, j As Long) As Variant
    ' Empty si esa ventana no tiene ninguna marca: asi la celda queda vacia
    ' en vez de mostrar un 0.00% que no es cierto.
    Dim k As String, arr As Variant
    If manVar Is Nothing Then Exit Function
    k = NormTxt(nom)
    If Not manVar.Exists(k) Then Exit Function
    arr = manVar(k)
    If CLng(arr(6 + j)) = 0 Then Exit Function
    VarVentana = Exp(CDbl(arr(j))) - 1
End Function





Private Function MasPrima(f As String, col As Long, aplica As Boolean) As String
    ' Toma una formula que ya viene como  =IFERROR(...,"")  y le suma la prima
    ' de Config C22. Si la de adentro devuelve "" el "" + numero revienta y el
    ' IFERROR de afuera deja la celda vacia: sin promedio no hay a que sumarle.
    Dim t As String
    MasPrima = f
    If Not aplica Then Exit Function
    If Len(f) = 0 Then Exit Function
    If Left$(f, 1) <> "=" Then Exit Function
    t = TerminoPrima(col)
    If Len(t) = 0 Then Exit Function
    MasPrima = "=IFERROR(" & Mid$(f, 2) & "+" & t & "," & Chr$(34) & Chr$(34) & ")"
End Function





Private Function TerminoPrima(col As Long) As String
    ' Config!$C$22 apunta a Config C22 y esta en puntos basicos: /10000 lo pasa a
    ' tasa. Config C23 decide si se reparte por dias o va parejo.
    Dim pb As String, cl As String

    pb = "Config!$C$22/10000"
    If UCase$(Left$(ModoPrima(), 4)) <> "PROR" Then
        TerminoPrima = pb
        Exit Function
    End If

    ' PRORRATEADO: los 100 pb son anuales.
    '   ventanas  -> por los dias de la propia ventana (fila fin menos base)
    '   Ann YTD   -> ya es anual, entra entera
    '   Mensual   -> la doceava parte
    If col >= CC_W1 And col <= CC_W1 + N_VENT - 1 Then
        cl = NumALetra(col)
        TerminoPrima = pb & "*(" & cl & "$" & CU_FIN & "-" & cl & "$" & CU_BASE & ")/365"
    ElseIf col = CC_ANN Then
        TerminoPrima = pb
    ElseIf col = CC_MEN Then
        TerminoPrima = pb & "/12"
    End If
End Function





Private Function ModoPrima() As String
    ' Config C23. Vacio o cualquier cosa que no empiece con PROR = PLANO.
    Dim modoCfg As String
    On Error Resume Next
    modoCfg = UCase$(Txt(ThisWorkbook.Worksheets(SH_CFG).Range("C23").Value))
    On Error GoTo 0
    If Len(modoCfg) = 0 Then modoCfg = "PLANO"
    ModoPrima = modoCfg
End Function





Private Function CatsDeConfig(celda As String) As Variant
    ' Lee una lista de categorias separada por | . Devuelve Empty si esta
    ' vacia, para que la fila simplemente no se escriba.
    Dim cad As String, p As Variant, res() As String, nn As Long
    On Error Resume Next
    cad = Txt(ThisWorkbook.Worksheets(SH_CFG).Range(celda).Value)
    On Error GoTo 0
    If Len(Trim$(cad)) = 0 Then Exit Function
    ReDim res(0 To 30)
    nn = 0
    For Each p In Split(cad, "|")
        If Len(Trim$(CStr(p))) > 0 Then
            If nn <= 30 Then
                res(nn) = Trim$(CStr(p))
                nn = nn + 1
            End If
        End If
    Next p
    If nn = 0 Then Exit Function
    ReDim Preserve res(0 To nn - 1)
    CatsDeConfig = res
End Function





Private Function EtiquetaNiveles() As String
    ' "Total PEN + USD en Niveles (C.O. USD 9.0%, PEN 10.0%)"
    ' Los niveles salen de Config C20 y C64 y son solo rotulo.
    Dim nPen As Double, nUsd As Double
    On Error Resume Next
    nPen = NumDef(ThisWorkbook.Worksheets(SH_CFG).Range("C20").Value, 0)
    nUsd = NumDef(ThisWorkbook.Worksheets(SH_CFG).Range("C21").Value, 0)
    On Error GoTo 0
    EtiquetaNiveles = "Total PEN + USD en Niveles (C.O. USD " & _
                      Format$(nUsd, "0.0%") & ", PEN " & _
                      Format$(nPen, "0.0%") & ")"
End Function





Private Function Pase(cMon As String, cl As String) As String
    ' Spread anual prorrateado por los dias de la ventana, solo si es USD.
    Pase = "IF(" & cMon & "=""USD"",Config!$C$15*(" & cl & "$" & CU_FIN & "-" & _
           cl & "$" & CU_BASE & ")/365,0)"
End Function





Private Function OrdenMonedas(ins As Object, orden As Variant) As Variant
    Dim d As Object, i As Long, arr As Variant, k As Variant
    Dim res() As String, n As Long
    Set d = NuevoDic()
    If IsArray(orden) Then
        If UBound(orden) >= LBound(orden) Then
            For i = LBound(orden) To UBound(orden)
                arr = ins(orden(i))
                If Not arr(4) Then d(CStr(arr(3))) = 1
            Next i
        End If
    End If
    If d.Count = 0 Then
        OrdenMonedas = Array()
        Exit Function
    End If
    ReDim res(0 To d.Count - 1)
    n = -1
    If d.Exists("PEN") Then n = n + 1: res(n) = "PEN"
    If d.Exists("USD") Then n = n + 1: res(n) = "USD"
    For Each k In d.Keys
        If CStr(k) <> "PEN" And CStr(k) <> "USD" Then n = n + 1: res(n) = CStr(k)
    Next k
    OrdenMonedas = res
End Function





Private Function OrdenCategorias(ins As Object, orden As Variant, mon As String) As Variant
    ' Sale en el orden de Config C17 (por riesgo). Lo que no este ahi, al final.
    Dim d As Object, i As Long, arr As Variant, k As Variant
    Dim res() As String, n As Long, p As Variant, ordCfg As String
    Set d = NuevoDic()
    For i = LBound(orden) To UBound(orden)
        arr = ins(orden(i))
        If Not arr(4) And CStr(arr(3)) = mon Then d(CStr(arr(2))) = 1
    Next i
    If d.Count = 0 Then
        OrdenCategorias = Array()
        Exit Function
    End If
    ReDim res(0 To d.Count - 1)
    n = -1
    ordCfg = Txt(ThisWorkbook.Worksheets(SH_CFG).Range("C17").Value)
    If Len(Trim$(ordCfg)) > 0 Then
        For Each p In Split(ordCfg, "|")
            For Each k In d.Keys
                If CLng(d(k)) = 1 And NormTxt(CStr(k)) = NormTxt(CStr(p)) Then
                    n = n + 1
                    res(n) = CStr(k)
                    d(k) = 2
                End If
            Next k
        Next p
    End If
    For Each k In d.Keys
        If CLng(d(k)) = 1 Then
            n = n + 1
            res(n) = CStr(k)
        End If
    Next k
    OrdenCategorias = res
End Function





Private Function OrdenFilas(ins As Object, orden As Variant, mon As String, _
                            cat As String, pesos As Object) As Variant
    ' Dentro de una categoria:
    '   - si esta en Config C60 Y al menos dos filas tienen vintage escrito,
    '     se ordena por VINTAGE, de mas antiguo a mas nuevo
    '   - en cualquier otro caso, por PESO DEL FONDO 1, de mayor a menor, y
    '     cuando el F1 empata desempata el PESO DEL FONDO 2, tambien de mayor
    '     a menor
    ' Los que no tienen el dato se van al final, en el orden de la hoja.
    Dim i As Long, n As Long, j As Long, k As Long, nVin As Long
    Dim arr As Variant, res() As String, cri() As Double, cri2() As Double
    Dim t As String, tv As Double, porVin As Boolean, antes As Boolean

    porVin = False
    If porVin Then
        ' Pedir orden por vintage y no haberlo escrito dejaba a todos empatados
        ' en 99999, o sea en el orden de la hoja Instrumentos, que parece "sin
        ' ordenar". Con menos de dos vintages se ordena por peso y listo.
        nVin = 0
        For i = LBound(orden) To UBound(orden)
            arr = ins(orden(i))
            If CStr(arr(3)) = mon And CStr(arr(2)) = cat And Not arr(4) Then
                If Num(arr(6)) > 0 Then nVin = nVin + 1
            End If
        Next i
        If nVin < 2 Then porVin = False
    End If
    n = -1
    ReDim res(0 To UBound(orden) - LBound(orden) + 1)
    ReDim cri(0 To UBound(orden) - LBound(orden) + 1)
    ReDim cri2(0 To UBound(orden) - LBound(orden) + 1)
    For i = LBound(orden) To UBound(orden)
        arr = ins(orden(i))
        If CStr(arr(3)) = mon And CStr(arr(2)) = cat And Not arr(4) Then
            n = n + 1
            res(n) = CStr(arr(0))
            If porVin Then
                cri(n) = Num(arr(6))
                If cri(n) <= 0 Then cri(n) = 99999
            Else
                ' peso al reves, para que el mayor quede primero con el mismo orden
                cri(n) = -PesoDe(pesos, CStr(arr(0)), 0)
            End If
            ' Desempate: F2, tambien de mayor a menor. Dos filas no pueden
            ' estar ordenadas por F1 y por F2 a la vez, asi que F1 manda y F2
            ' solo decide cuando F1 empata (tipico: las dos en cero).
            cri2(n) = -PesoDe(pesos, CStr(arr(0)), 1)
        End If
    Next i
    If n < 0 Then
        OrdenFilas = Array()
        Exit Function
    End If
    ReDim Preserve res(0 To n)
    ReDim Preserve cri(0 To n)
    ReDim Preserve cri2(0 To n)
    For j = 0 To n - 1
        For k = j + 1 To n
            antes = False
            If cri(k) < cri(j) Then
                antes = True
            ElseIf cri(k) = cri(j) Then
                If cri2(k) < cri2(j) Then antes = True
            End If
            If antes Then
                tv = cri(j): cri(j) = cri(k): cri(k) = tv
                tv = cri2(j): cri2(j) = cri2(k): cri2(k) = tv
                t = res(j): res(j) = res(k): res(k) = t
            End If
        Next k
    Next j
    OrdenFilas = res
End Function





Private Function PesoDe(pesos As Object, cod As String, Optional idx As Long = 0) As Double
    ' idx 0 = Fondo 1, idx 1 = Fondo 2.
    Dim a As Variant
    If pesos Is Nothing Then Exit Function
    If Not pesos.Exists(cod) Then Exit Function
    a = pesos(cod)
    If Not IsArray(a) Then
        If idx = 0 Then PesoDe = Num(a)
        Exit Function
    End If
    If idx >= LBound(a) And idx <= UBound(a) Then PesoDe = Num(a(idx))
End Function





Private Function PesosPorCodigo() As Object
    ' codigo -> Array(peso F1, peso F2), leido de BD_Pesos (columnas 7 y 8).
    ' F1 es el que ordena las filas dentro de cada categoria, de mayor a menor;
    ' F2 solo entra a desempatar.
    Dim d As Object, ws As Worksheet, i As Long, ult As Long, cod As String
    Dim a As Variant
    Set d = NuevoDic()
    Set PesosPorCodigo = d
    If Not HojaExiste(SH_PES) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_PES)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
        If Len(cod) > 0 Then
            ' Se SUMA, no se pisa. El cuadro muestra SUMIFS sobre el codigo, asi
            ' que si un codigo apareciera dos veces en BD_Pesos el orden tiene
            ' que usar la misma suma o el cuadro saldria "desordenado".
            If d.Exists(cod) Then
                a = d(cod)
                d(cod) = Array(Num(a(0)) + Num(ws.Cells(i, 7).Value), _
                               Num(a(1)) + Num(ws.Cells(i, 8).Value))
            Else
                d(cod) = Array(Num(ws.Cells(i, 7).Value), Num(ws.Cells(i, 8).Value))
            End If
        End If
    Next i
End Function




Private Sub Escala(rng As Range)
    ' divergente y simetrica en cero: el 0.00% queda blanco, no rojo
    Dim fc As ColorScale, c As Range
    Dim vals() As Double, n As Long, i As Long, k As Long, t As Double, M As Double

    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    rng.FormatConditions.Delete
    On Error GoTo 0

    ReDim vals(1 To rng.Cells.Count)
    n = 0
    For Each c In rng.Cells
        If Not IsError(c.Value) Then
            If IsNumeric(c.Value) And Len(Txt(c.Text)) > 0 Then
                n = n + 1
                vals(n) = Abs(CDbl(c.Value))
            End If
        End If
    Next c
    If n = 0 Then Exit Sub

    For i = 1 To n - 1
        For k = i + 1 To n
            If vals(k) < vals(i) Then
                t = vals(i): vals(i) = vals(k): vals(k) = t
            End If
        Next k
    Next i
    i = CLng(Int(0.9 * n))
    If i < 1 Then i = 1
    If i > n Then i = n
    M = vals(i)
    If M <= 0 Then M = vals(n)
    If M <= 0 Then Exit Sub

    Set fc = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
    With fc.ColorScaleCriteria(1)
        .Type = xlConditionValueNumber
        .Value = -M
        .FormatColor.Color = RGB(230, 124, 115)
    End With
    With fc.ColorScaleCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .FormatColor.Color = RGB(255, 255, 255)
    End With
    With fc.ColorScaleCriteria(3)
        .Type = xlConditionValueNumber
        .Value = M
        .FormatColor.Color = RGB(125, 191, 138)
    End With
End Sub



Private Function LeerFondos() As Object
    ' codigo SBS -> Array(curr yield, WTD, 5D, MTD, MayoTD, YTD, FY)
    Dim ws As Worksheet, d As Object, i As Long, ult As Long, j As Long
    Dim cod As String, v(0 To 6) As Variant, hay As Boolean

    Set d = NuevoDic()
    Set LeerFondos = d
    If Not HojaExiste(SH_FON) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_FON)
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Function

    For i = 2 To ult
        cod = SinGuiones(Txt(ws.Cells(i, 1).Value))
        If Len(cod) > 0 Then
            hay = False
            For j = 0 To 6
                If EsNumero(ws.Cells(i, j + 3).Value) Then
                    v(j) = CDbl(ws.Cells(i, j + 3).Value)
                    If j > 0 Then hay = True
                Else
                    v(j) = Empty
                End If
            Next j
            If hay Then
                d(cod) = Array(v(0), v(1), v(2), v(3), v(4), v(5), v(6))
            End If
        End If
    Next i
End Function




Private Function ArmaNombre(base As String, cupon As Double, venc As Double) As String
    ' Solo pega lo que existe. Un fondo no tiene cupon ni vencimiento, asi que
    ' su nombre sale tal cual lo escribiste, sin nada detras.
    Dim out As String
    out = Trim$(base)
    If cupon > 0 Then out = out & " " & Format$(cupon, "0.00")
    If venc >= 10000 Then out = out & " " & Format$(CDate(venc), "dd/mm/yy")
    ArmaNombre = out
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





