// ============================================================================
//  Consulta "Maestro"  -  Base_Vector.xlsx
//  Ultimo dato de cada codigo (la fila de la ultima fecha en Precios) mas lo
//  que escribes a mano en la tabla Codigos (Nombre, ISIN, Spread pb).
// ============================================================================
let
    Fuente       = Precios,
    Agrupado     = Table.Group(Fuente, {"CodigoSBS"}, {{"Ultima", each Table.Max(_, "Fecha"), type record}}),
    Expandido    = Table.ExpandRecordColumn(Agrupado, "Ultima",
                       {"Fecha", "Nemonico", "Emisor", "Moneda", "Cupon", "Vencimiento", "Rating", "Duracion", "PrecioLimpio"},
                       {"UltimaFecha", "Nemonico", "Emisor", "Moneda", "Cupon", "Vencimiento", "Rating", "Duracion", "UltimoPrecio"}),

    TablaCodigos = Excel.CurrentWorkbook(){[Name = "Codigos"]}[Content],
    CodLimpio    = Table.AddColumn(TablaCodigos, "CodigoSBS", each
                       Text.Replace(Text.Replace(Text.From([Codigo SBS]), "-", ""), " ", ""), type text),
    Unido        = Table.NestedJoin(Expandido, {"CodigoSBS"}, CodLimpio, {"CodigoSBS"}, "C", JoinKind.LeftOuter),
    ConManual    = Table.ExpandTableColumn(Unido, "C", {"Nombre", "ISIN", "Spread (pb)"}, {"Nombre", "ISIN", "Spread"}),
    Final        = Table.SelectColumns(ConManual,
                       {"CodigoSBS", "Nombre", "UltimaFecha", "Nemonico", "Emisor", "Moneda", "Cupon",
                        "Vencimiento", "Rating", "Duracion", "UltimoPrecio", "ISIN", "Spread"})
in
    Final
