using System;
using System.Collections.Generic;

namespace XNSeguimentCompres.Domain
{
    /// <summary>
    /// Validaciones de negocio para Seguimiento de Compras.
    /// 
    /// ❗ No depende de SAP ni UI.
    /// ✅ Puede ser usada desde AddOn, API o App móvil.
    /// </summary>
    public class SegComValidator
    {
        /// <summary>
        /// Valida toda la información antes de guardar.
        /// </summary>
        public bool Validate(SegComHeader h, List<SegComLine> lines, out string error)
        {
            // ---- Cabecera ----
            if (string.IsNullOrWhiteSpace(h.CardCode))
            {
                error = "El proveïdor és obligatori.";
                return false;
            }

            // ---- Líneas ----
            if (lines == null || lines.Count == 0)
            {
                error = "Cal afegir com a mínim una línia de seguiment.";
                return false;
            }

            foreach (var l in lines)
            {
                if (string.IsNullOrWhiteSpace(l.Dscription))
                {
                    error = "Hi ha línies sense descripció.";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(l.LineStatus))
                {
                    error = "Cal seleccionar un estat a cada línia.";
                    return false;
                }

                if (l.LineOrder <= 0)
                {
                    error = "El número d'ordre és incorrecte.";
                    return false;
                }
            }

            

            // ---- OK ----
            error = null;
            return true;
        }
    }
}



