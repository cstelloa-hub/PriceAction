// ============================================================================
//  Consulta "Precios"  -  Base_Vector.xlsx
//  Lee todos los vectores de precios de la SBS de una carpeta (archivos
//  "AAAAMMDD RFL.xls"), se queda con los codigos de la tabla Codigos y deja
//  una fila por codigo y dia, ordenada por fecha. Se actualiza con Refresh.
//
//  Columnas del vector (mismas letras que Config de Base_RF):
//     A codigo (con guiones)   C nemonico   E emisor   F moneda
//     J precio limpio          M interes    N TIR      Q vencimiento
//     R cupon                  U rating     X duracion
//     Spread: la letra la pones en ColSpread (en el vector SBS suele venir mas
//     a la derecha; mira la cabecera del archivo). "" = el vector no lo trae.
// ============================================================================
let
    // ---------------- PARAMETROS (edita estas dos lineas) ----------------
    RutaCarpeta  = "\\SERVIDOR\carpeta\de\vectores\",
    FechaInicio  = #date(2025, 10, 31),
    ColSpread    = "",            // letra de la columna del spread en el vector, p. ej. "Z". "" = no hay
    FactorPb     = 100,           // el spread del vector viene en % (0.0185 = 185 pb)? pon 10000. En pb? pon 1. En 1.85? pon 100

    // letra de columna -> numero
    LetraANum = (s as text) as number =>
        List.Accumulate(Text.ToList(Text.Upper(Text.Trim(s))), 0, (acc, ch) => acc * 26 + (Character.ToNumber(ch) - 64)),
    NumSpread = if ColSpread = null or Text.Trim(ColSpread) = "" then 0 else LetraANum(ColSpread),

    // ---------------- codigos a seguir: tabla "Codigos" de este libro ----
    TablaCodigos = Excel.CurrentWorkbook(){[Name = "Codigos"]}[Content],
    ListaCodigos = List.Buffer(
        List.Transform(
            List.RemoveNulls(TablaCodigos[Codigo SBS]),
            each Text.Replace(Text.Replace(Text.From(_), "-", ""), " ", ""))),

    // ---------------- archivos de la carpeta ----------------------------
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

    // ---------------- leer un vector: columnas por posicion ---------------
    LeerUno = (contenido as binary) as table =>
        let
            Libro = Excel.Workbook(contenido, false),
            Hoja  = Libro{0}[Data],
            Cols  = Table.ColumnNames(Hoja),
            Col   = (n as number) => if n <= List.Count(Cols) then Table.Column(Hoja, Cols{n - 1}) else List.Repeat({null}, Table.RowCount(Hoja)),
            Sel   = Table.FromColumns(
                        {Col(1), Col(3), Col(5), Col(6), Col(10), Col(13), Col(14), Col(17), Col(18), Col(21), Col(24),
                         if NumSpread > 0 then Col(NumSpread) else List.Repeat({null}, Table.RowCount(Hoja))},
                        {"Codigo", "Nemonico", "Emisor", "Moneda", "PrecioLimpio", "Interes", "TIR", "Vencimiento", "Cupon", "Rating", "Duracion", "Spread"})
        in
            Sel,

    Leidos     = Table.AddColumn(EnRango, "Datos", each try LeerUno([Content]) otherwise null),
    SinError   = Table.SelectRows(Leidos, each [Datos] <> null),
    Expandido  = Table.ExpandTableColumn(
                     Table.SelectColumns(SinError, {"Fecha", "Datos"}), "Datos",
                     {"Codigo", "Nemonico", "Emisor", "Moneda", "PrecioLimpio", "Interes", "TIR", "Vencimiento", "Cupon", "Rating", "Duracion", "Spread"}),

    // ---------------- solo los codigos seguidos ---------------------------
    CodLimpio  = Table.AddColumn(Expandido, "CodigoSBS", each
                     if [Codigo] = null then null
                     else Text.Replace(Text.Replace(Text.From([Codigo]), "-", ""), " ", ""), type text),
    Seguidos   = Table.SelectRows(CodLimpio, each [CodigoSBS] <> null and List.Contains(ListaCodigos, [CodigoSBS])),

    // ---------------- tipos y moneda normalizada --------------------------
    Tipos      = Table.TransformColumnTypes(Seguidos, {
                     {"PrecioLimpio", type number}, {"Interes", type number}, {"TIR", type number},
                     {"Cupon", type number}, {"Duracion", type number}, {"Spread", type number}}),
    SpreadPb   = Table.TransformColumns(Tipos, {{"Spread", each if _ = null then null else _ * FactorPb, type number}}),
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
