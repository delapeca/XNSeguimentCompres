using System;

namespace XNSeguimentCompres.Domain
{
    /// <summary>
    /// Representa una línea de seguimiento.
    /// Independiente de SAP UI.
    /// </summary>
    public class SegComLine
    {
        public int LineId { get; set; }
        public int LineOrder { get; set; }
        public string Dscription { get; set; }
        public System.DateTime? Date { get; set; }
        public string Hour { get; set; }
        public string LineStatus { get; set; }
        public DateTime? StatusDate { get; set; }

    }
}

