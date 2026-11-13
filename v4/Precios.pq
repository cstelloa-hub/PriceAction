// ============================================================================
//  Consulta "Precios"  -  precios del vector SBS dentro del libro
//  Lee todos los vectores de una carpeta (archivos "AAAAMMDD RFL.xls"), se
//  queda con los codigos de la tabla Codigos y deja una fila por codigo y dia,
//  ordenada por fecha. Se actualiza con Refresh.
//
//  OJO: los "RFL.xls" de la SBS son una tabla HTML con extension .xls (la
//  cabecera trae "C&Oacute;DIGO"). Se leen como HTML; si algun archivo fuera
//  Excel de verdad o texto con tabs, tambien se lee.
//
//  Columnas del vector (posicion):
//     1 codigo (con guiones)  3 nemonico  5 emisor  6 moneda  10 precio limpio %
//    13 interes  14 TIR  17 vencimiento  18 cupon  21 rating  24 duracion
//    spread: la letra la pones en ColSpread (en tu vector es la O)
// ============================================================================
let
    // ---------------- PARAMETROS (edita estas lineas) ----------------------
    RutaCarpeta  = "Z:\10. Varios\Jose Carlos Cortes\9. Vector\Vectores de precios\RF\",
    FechaInicio  = #date(2025, 10, 31),
    ColSpread    = "O",           // letra de la columna del spread en el vector. "" = no hay
    FactorPb     = 100,           // spread del vector x FactorPb = puntos basicos (0.3165 -> 31.65 pb con 100)

    // ---------------- utilidades -------------------------------------------
    LetraANum = (s as text) as number =>
        List.Accumulate(Text.ToList(Text.Upper(Text.Trim(s))), 0, (acc, ch) => acc * 26 + (Character.ToNumber(ch) - 64)),
    NumSpread = if ColSpread = null or Text.Trim(ColSpread) = "" then 0 else LetraANum(ColSpread),
    Nombres   = {"Codigo", "Nemonico", "Emisor", "Moneda", "PrecioLimpio", "Interes", "TIR", "Vencimiento", "Cupon", "Rating", "Duracion", "Spread"},
    Posic     = {1, 3, 5, 6, 10, 13, 14, 17, 18, 21, 24, NumSpread},

    // ---------------- codigos a seguir: tabla "Codigos" de este libro ------
    TablaCodigos = Excel.CurrentWorkbook(){[Name = "Codigos"]}[Content],
    ListaCodigos = List.Buffer(
        List.Transform(
            List.RemoveNulls(TablaCodigos[Codigo SBS]),
            each Text.Replace(Text.Replace(Text.From(_), "-", ""), " ", ""))),

    // ---------------- archivos de la carpeta -------------------------------
    Archivos   = Folder.Files(RutaCarpeta),
    SoloRFL    = Table.SelectRows(Archivos, each
                     Text.EndsWith(Text.Upper([Name]), "RFL.XLS") or
                     Text.EndsWith(Text.Upper([Name]), "RFL.XLSX")),
    ConFecha   = Table.AddColumn(SoloRFL, "Fecha", each
                     try #date(Number.FromText(Text.Start([Name], 4)),
                               Number.FromText(Text.Middle([Name], 4, 2)),
                               Number.FromText(Text.Middle([Name], 6, 2)))
                     otherwise null, type date),
    EnRango    = Table.SelectRows(ConFecha, each [Fecha] <> null and [Fecha] >= FechaInicio),

    // ---------------- leer un vector ---------------------------------------
    LeerUno = (contenido as binary) as table =>
        let
            Inicio   = try Text.Lower(Text.Start(Text.FromBinary(contenido, 1252), 3000)) otherwise "",
            EsHtml   = Text.Contains(Inicio, "<table") or Text.Contains(Inicio, "<html") or Text.Contains(Inicio, "<tr"),
            // HTML: cada fila <tr>, cada columna por su posicion <td>
            DesdeHtml = () =>
                Html.Table(contenido,
                    List.Transform(List.Positions(Nombres), (i) =>
                        {Nombres{i}, if Posic{i} > 0 then "td:nth-child(" & Text.From(Posic{i}) & ")" else "td:nth-child(9999)"}),
                    [RowSelector = "table tr"]),
            // Excel o texto: columnas por posicion
            DesdeTabla = (Hoja as table) =>
                let
                    Cols = Table.ColumnNames(Hoja),
                    Col  = (n as number) => if n > 0 and n <= List.Count(Cols) then Table.Column(Hoja, Cols{n - 1})
                                            else List.Repeat({null}, Table.RowCount(Hoja))
                in
                    Table.FromColumns(List.Transform(Posic, each Col(_)), Nombres),
            ComoExcel = if EsHtml then null else try Excel.Workbook(contenido, false){0}[Data] otherwise null,
            Resultado = if EsHtml then DesdeHtml()
                        else if ComoExcel <> null then DesdeTabla(ComoExcel)
                        else DesdeTabla(Csv.Document(contenido, [Delimiter = "#(tab)", Encoding = 1252, QuoteStyle = QuoteStyle.None]))
        in
            Resultado,

    Leidos     = Table.AddColumn(EnRango, "Datos", each try LeerUno([Content]) otherwise null),
    SinError   = Table.SelectRows(Leidos, each [Datos] <> null),
    Expandido  = Table.ExpandTableColumn(Table.SelectColumns(SinError, {"Fecha", "Datos"}), "Datos", Nombres),

    // ---------------- solo los codigos seguidos ----------------------------
    CodLimpio  = Table.AddColumn(Expandido, "CodigoSBS", each
                     if [Codigo] = null then null
                     else Text.Replace(Text.Replace(Text.Trim(Text.From([Codigo])), "-", ""), " ", ""), type text),
    Seguidos   = Table.SelectRows(CodLimpio, each [CodigoSBS] <> null and List.Contains(ListaCodigos, [CodigoSBS])),

    // ---------------- numeros (punto decimal, como vienen en el HTML) ------
    ANumero    = (x) => if x = null then null
                        else if Value.Is(x, type number) then x
                        else try Number.FromText(Text.Replace(Text.Trim(Text.From(x)), ",", ""), "en-US") otherwise null,
    Numeros    = Table.TransformColumns(Seguidos, {
                     {"PrecioLimpio", ANumero, type number}, {"Interes", ANumero, type number},
                     {"TIR", ANumero, type number}, {"Cupon", ANumero, type number},
                     {"Duracion", ANumero, type number}, {"Spread", ANumero, type number}}),
    SpreadPb   = Table.TransformColumns(Numeros, {{"Spread", each if _ = null then null else _ * FactorPb, type number}}),
    MonedaN    = Table.TransformColumns(SpreadPb, {{"Moneda", each
                     let m = Text.Upper(Text.Remove(Text.From(_), {".", "/", " "}))
                     in  if m = "S" or m = "PEN" or m = "SOLES" or m = "NUEVOSSOLES" then "PEN"
                         else if m = "US$" or m = "USD" or m = "US" or m = "DOLARES" then "USD"
                         else Text.From(_), type text}}),
    Final      = Table.Sort(
                     Table.SelectColumns(MonedaN, {"Fecha", "CodigoSBS", "Nemonico", "Emisor", "Moneda",
                                                   "PrecioLimpio", "Interes", "Cupon", "Vencimiento", "TIR", "Duracion", "Spread", "Rating"}),
                     {{"Fecha", Order.Ascending}, {"CodigoSBS", Order.Ascending}})
in
    Final
