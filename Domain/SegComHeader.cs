namespace XNSeguimentCompres.Domain
{
    /// <summary>
    /// Representa la cabecera de un Seguimiento de Pedido de Compra.
    /// Independiente de SAP UI → apto para API o app móvil.
    /// </summary>
    public class SegComHeader
    {
        public int DocEntry { get; set; }
        public int DocNum { get; set; }
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string NumAtCard { get; set; }
        public int Status { get; set; } = 0;
        public int BaseEntry { get; set; }   // OPOR.DocEntry
        public int BaseNum { get; set; }   // OPOR.DocNum
        public System.DateTime DocDate { get; set; }
    }
}

