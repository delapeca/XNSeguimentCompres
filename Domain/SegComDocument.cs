using System.Collections.Generic;

namespace XNSeguimentCompres.Domain
{
    /// <summary>
    /// Representa un seguiment complet (capçalera + línies)
    /// per simplificar la lectura i transport de dades.
    /// </summary>
    public class SegComDocument
    {
        public SegComHeader Header { get; set; }
        public List<SegComLine> Lines { get; set; }

        public SegComDocument()
        {
            Lines = new List<SegComLine>();
        }
    }
}

