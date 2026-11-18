// ============================================================================
//  Consulta "Marcas"  -  marcas de los fondos tradicionales dentro del libro
//  Lee el consolidado de Marcas (un solo libro, hoja "3"), se queda con los
//  codigos de la tabla Codigos y deja UNA FILA POR MARCA:
//     Fecha . CodigoSBS . Nombre . VarAjustada . LnVar
//  LnVar = LN(1 + VarAjustada): con eso el retorno de cualquier ventana es
//     EXP( SUMA de LnVar entre base y fin ) - 1   (una sola SUMIFS)
//  Se actualiza con Refresh, igual que Precios. Marcas se lee de disco, no se
//  escribe nunca.
//
//  Las columnas se buscan por el TEXTO de la cabecera (Dia / Fecha, Codigo
//  SBS, Var Ajustada, Nombre); si cambian de sitio no pasa nada. Si no las
//  encuentra, el error te lista las cabeceras que vio.
// ============================================================================
let
    // ---------------- PARAMETROS (edita estas lineas) ----------------------
    RutaMarcas   = "Z:\RUTA\DEL\LIBRO\Marcas.xlsx",   // ruta completa, con extension
    HojaMarcas   = "3",                                // nombre de la pestana tal cual
    FilaCabecera = 3,                                  // fila donde estan los titulos
    FechaInicio  = #date(2024, 12, 31),                // marcas anteriores se descartan

    // ---------------- utilidades -------------------------------------------
    ACodigo = (x) as text => if x = null then "" else Text.Upper(Text.Remove(Text.Trim(Text.From(x)), {"-", " "})),
    ANumero = (x) => if x = null then null
                     else if Value.Is(x, type number) then x
                     else try Number.FromText(Text.Replace(Text.Trim(Text.From(x)), ",", "."), "en-US") otherwise null,
    AFecha  = (x) => if x = null then null
                     else if Value.Is(x, type date) then x
                     else if Value.Is(x, type datetime) then Date.From(x)
                     else if Value.Is(x, type number) then (try Date.From(x) otherwise null)
                     else try Date.From(DateTime.FromText(Text.Trim(Text.From(x)), [Culture = "es-PE"]))
                          otherwise (try Date.From(DateTime.FromText(Text.Trim(Text.From(x)), [Culture = "en-US"])) otherwise null),

    // ---------------- codigos a seguir: tabla "Codigos" de este libro ------
    TablaCodigos = Excel.CurrentWorkbook(){[Name = "Codigos"]}[Content],
    ListaCodigos = List.Buffer(List.Distinct(List.Transform(List.RemoveNulls(TablaCodigos[Codigo SBS]), each ACodigo(_)))),

    // ---------------- abrir Marcas -----------------------------------------
    Libro   = Excel.Workbook(File.Contents(RutaMarcas), null, true),
    Hojas   = Table.SelectRows(Libro, each [Kind] = "Sheet"),
    Hoja    = if Table.IsEmpty(Table.SelectRows(Hojas, each [Name] = HojaMarcas))
              then error "No encuentro la hoja '" & HojaMarcas & "' en Marcas. Hojas: " & Text.Combine(Hojas[Name], " | ")
              else Hojas{[Name = HojaMarcas]}[Data],
    SinTop  = Table.Skip(Hoja, FilaCabecera - 1),
    Cab     = Table.PromoteHeaders(SinTop, [PromoteAllScalars = true]),
    Nombres = Table.ColumnNames(Cab),

    // columna cuya cabecera contenga el patron; los patrones van en orden de preferencia
    Busca = (patrones as list, obligatoria as logical) =>
        let
            PorPatron = List.Transform(patrones, (p) => List.Select(Nombres, (n) => Text.Contains(Text.Upper(n), p))),
            Hallazgos = List.Select(PorPatron, each not List.IsEmpty(_))
        in
            if not List.IsEmpty(Hallazgos) then Hallazgos{0}{0}
            else if obligatoria then error "No encuentro la columna " & Text.Combine(patrones, " / ") &
                                            " en la fila " & Text.From(FilaCabecera) & ". Cabeceras: " & Text.Combine(Nombres, " | ")
            else null,
    ColDia = Busca({"DIA", "DÍA", "FECHA"}, true),
    ColCod = Busca({"CODIGO SBS", "CÓDIGO SBS", "COD SBS", "CODIGO", "CÓDIGO"}, true),
    ColVar = Busca({"VAR AJ", "VARIACION AJ", "VARIACIÓN AJ", "AJUSTADA", "VAR"}, true),
    ColNom = Busca({"NOMBRE", "FONDO"}, false),

    // ---------------- quedarse con las cuatro columnas ---------------------
    Sel     = Table.SelectColumns(Cab, List.Distinct({ColDia, ColCod, ColVar} & (if ColNom = null then {} else {ColNom}))),
    Ren     = Table.RenameColumns(Sel, {{ColDia, "Fecha"}, {ColCod, "CodigoSBS"}, {ColVar, "VarAjustada"}} &
                                        (if ColNom = null or ColNom = ColCod then {} else {{ColNom, "Nombre"}})),
    ConNom  = if Table.HasColumns(Ren, "Nombre") then Ren else Table.AddColumn(Ren, "Nombre", each null, type text),
    Tipos   = Table.TransformColumns(ConNom, {
                  {"Fecha", AFecha, type date},
                  {"CodigoSBS", ACodigo, type text},
                  {"VarAjustada", ANumero, type number},
                  {"Nombre", each if _ = null then null else Text.From(_), type text}}),

    // ---------------- filas validas y codigos seguidos ---------------------
    //  VarAjustada <= -1 (perdida total o basura) se descarta: LN no existe.
    Validas  = Table.SelectRows(Tipos, each [Fecha] <> null and [Fecha] >= FechaInicio
                                         and [CodigoSBS] <> "" and [VarAjustada] <> null and [VarAjustada] > -1),
    Seguidos = if List.IsEmpty(ListaCodigos) then Validas
               else Table.SelectRows(Validas, each List.Contains(ListaCodigos, [CodigoSBS])),

    // ---------------- LN(1+var) y orden -------------------------------------
    ConLn    = Table.AddColumn(Seguidos, "LnVar", each Number.Ln(1 + [VarAjustada]), type number),
    Final    = Table.Sort(Table.SelectColumns(ConLn, {"Fecha", "CodigoSBS", "Nombre", "VarAjustada", "LnVar"}),
                          {{"CodigoSBS", Order.Ascending}, {"Fecha", Order.Ascending}})
in
    Final
