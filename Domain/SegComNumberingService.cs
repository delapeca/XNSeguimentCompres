using SAPbobsCOM;

namespace XNSeguimentCompres.Domain
{
    /// <summary>
    /// Servicio de numeración de documentos de Seguimiento.
    /// Permite obtener el próximo número de documento desde los UDT
    /// o desde la tabla de series NNM1.
    /// 
    /// NO contiene lógica de UI.
    /// </summary>
    public class SegComNumberingService
    {
        private readonly Company _company;

        public SegComNumberingService(Company company)
        {
            _company = company;
        }

        /// <summary>
        /// Obtiene el siguiente número de documento desde el UDT @XNSEGCOM.
        /// </summary>
        public int GetNextDocNumFromUdt()
        {
            Recordset rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(@"SELECT ISNULL(MAX(CAST(DocNum AS INT)), 0) AS MaxNum FROM ""@XNSEGCOM""");

            int max = (int)rs.Fields.Item("MaxNum").Value;
            return max + 1;
        }
    }
}

