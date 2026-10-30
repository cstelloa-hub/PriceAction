Attribute VB_Name = "mod_Guardar"
Option Explicit

'==============================================================================
' mod_Guardar  v1.2  (para PriceAction v3: Config C5 corte, C25 carpeta; copia hojas ocultas)
'
' Modulo aparte. No toca nada de mod_Iliquidos: ni lo lee, ni lo llama, ni
' depende de sus funciones. Se puede pegar y borrar sin romper el otro.
'
' Que hace:
'   Guarda una FOTO del CUADRO y del Retornos en un .xlsx nuevo, con la fecha
'   de corte en el nombre. La foto va en valores: sin formulas, sin nombres
'   definidos, sin vinculos al libro original. Se abre en cualquier PC.
'
' Macro publica:
'   GuardarCuadro      guarda la foto.
'
' Donde guarda:
'   Config C25 si tiene una carpeta escrita.
'   Si esta vacia, en <carpeta del libro>\Resumenes.
'   Nombre: Iliquidos_CUADRO_<corte yyyymmdd>.xlsx  (_v2, _v3 si ya existe)
'
' El corte lo lee de Config C5. Si no hay fecha ahi, usa la ultima de Precios.
'==============================================================================

Private Const G_CFG As String = "Config"
Private Const G_CUA As String = "CUADRO"
Private Const G_RET As String = "Retornos"
Private Const G_FON As String = "Fondos"


Public Sub GuardarCuadro()
    Dim wbNew As Workbook, ws As Worksheet
    Dim carpeta As String, base As String, ruta As String
    Dim dCorte As Date, n As Long
    Dim scrPrev As Boolean, evPrev As Boolean, alPrev As Boolean
    Dim t0 As Single, msg As String

    On Error GoTo Falla
    t0 = Timer
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    alPrev = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If Not HojaExiste(G_CUA) Then
        Err.Raise vbObjectError + 101, , _
            "No encuentro la hoja '" & G_CUA & "'. Corre Recalcular primero."
    End If

    ' ---- fecha de corte, solo para el nombre del archivo
    dCorte = CorteDeConfig()

    ' ---- carpeta destino
    carpeta = SinBarraFinal(Texto(ThisWorkbook.Worksheets(G_CFG).Range("C25").Value))
    If Len(carpeta) = 0 Then
        carpeta = ThisWorkbook.Path & Application.PathSeparator & "Resumenes"
    End If
    If Len(Dir(carpeta, vbDirectory)) = 0 Then MkDir carpeta

    ' ---- nombre sin pisar nada
    base = carpeta & Application.PathSeparator & _
           "Iliquidos_CUADRO_" & Format$(dCorte, "yyyymmdd")
    ruta = base & ".xlsx"
    n = 1
    Do While Len(Dir(ruta)) > 0
        n = n + 1
        ruta = base & "_v" & n & ".xlsx"
    Loop

    ' ---- copiar las hojas a un libro nuevo. Retornos y Fondos viven ocultas
    '      en PriceAction: se muestran un momento para poder copiarlas.
    Dim visRet As XlSheetVisibility, visFon As XlSheetVisibility, hojas As Variant
    hojas = Array(G_CUA)
    If HojaExiste(G_RET) And HojaExiste(G_FON) Then
        visRet = ThisWorkbook.Worksheets(G_RET).Visible
        visFon = ThisWorkbook.Worksheets(G_FON).Visible
        ThisWorkbook.Worksheets(G_RET).Visible = xlSheetVisible
        ThisWorkbook.Worksheets(G_FON).Visible = xlSheetVisible
        hojas = Array(G_CUA, G_RET, G_FON)
    ElseIf HojaExiste(G_RET) Then
        visRet = ThisWorkbook.Worksheets(G_RET).Visible
        ThisWorkbook.Worksheets(G_RET).Visible = xlSheetVisible
        hojas = Array(G_CUA, G_RET)
    End If
    ThisWorkbook.Worksheets(hojas).Copy
    Set wbNew = ActiveWorkbook
    On Error Resume Next
    If HojaExiste(G_RET) Then ThisWorkbook.Worksheets(G_RET).Visible = visRet
    If HojaExiste(G_FON) Then ThisWorkbook.Worksheets(G_FON).Visible = visFon
    On Error GoTo Falla

    ' ---- congelar en valores y cortar cualquier cordon umbilical
    For Each ws In wbNew.Worksheets
        AValores ws
    Next ws
    BorrarNombres wbNew

    Application.DisplayAlerts = False
    wbNew.SaveAs Filename:=ruta, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

    Application.DisplayAlerts = alPrev
    Application.EnableEvents = evPrev
    Application.ScreenUpdating = scrPrev

    msg = "Foto guardada en " & Format$(Timer - t0, "0.0") & " s." & vbCrLf & vbCrLf
    msg = msg & ruta & vbCrLf & vbCrLf
    msg = msg & "Corte: " & Format$(dCorte, "dd/mm/yyyy") & vbCrLf
    msg = msg & "Hojas: " & G_CUA & IIf(HojaExiste(G_RET), ", " & G_RET, "") & IIf(HojaExiste(G_FON), ", " & G_FON, "") & vbCrLf
    msg = msg & "Va en valores, sin formulas ni vinculos a este libro."
    MsgBox msg, vbInformation, "GuardarCuadro"
    Exit Sub

Falla:
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "GuardarCuadro se detuvo." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "GuardarCuadro"
End Sub


'------------------------------------------------------------------- auxiliares

Private Sub AValores(ws As Worksheet)
    ' Reemplaza formulas por su resultado, solo en el rango usado. Hacerlo
    ' sobre ws.Cells completo demora una eternidad: son un millon de filas.
    Dim rg As Range
    On Error Resume Next
    Set rg = ws.UsedRange
    On Error GoTo 0
    If rg Is Nothing Then Exit Sub
    rg.Value = rg.Value
End Sub


Private Sub BorrarNombres(wb As Workbook)
    ' Los nombres definidos (prCod, reSuma, cfgSpread...) apuntan al libro
    ' original. En la foto no sirven de nada y ensucian con #REF.
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
        v = ThisWorkbook.Worksheets(G_CFG).Range("C5").Value
        If IsDate(v) Then
            CorteDeConfig = CDate(v)
            Exit Function
        End If
    End If
    If HojaExiste("Precios") Then
        Set ws = ThisWorkbook.Worksheets("Precios")
        ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ult >= 2 Then CorteDeConfig = CDate(Application.WorksheetFunction.Max(ws.Range("A2:A" & ult)))
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
