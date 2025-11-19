using System;
using System.Linq;
using System.Collections.Generic;
using SAPbobsCOM;
using XNSeguimentCompres.Domain;
using System.Diagnostics;

namespace XNSeguimentCompres.Data
{
    /// <summary>
    /// Se encarga de persistir los datos de Seguimiento de Compras en los UDT:
    /// @XNSEGCOM (cabecera) y @XNSEGCOM01 (líneas)
    /// 
    /// Importante → No contiene lógica de validación ni UI.
    /// Solo ejecuta operaciones de insertar / actualizar.
    /// </summary>
    public class SegComRepository
    {
        private readonly Company _company;

        public SegComRepository(Company company)
        {
            _company = company;
        }

        // SegComRepository.cs
        private string GetChildTableNameFromDb(string udoCode)
        {
            var rs = (SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // Primer intent: UDO1 (molt habitual)
            try
            {
                rs.DoQuery($@"SELECT TOP 1 TableName FROM UDO1 WHERE Code = '{udoCode}' ORDER BY LineId");
                if (rs.RecordCount > 0)
                    return rs.Fields.Item("TableName").Value.ToString().Trim();
            }
            catch { /* ignorem */ }

            // Segon intent: OUDC (algunes instal·lacions)
            try
            {
                rs.DoQuery($@"SELECT TOP 1 ChildTbl AS TableName FROM OUDC WHERE Code = '{udoCode}' ORDER BY LineId");
                if (rs.RecordCount > 0)
                    return rs.Fields.Item("TableName").Value.ToString().Trim();
            }
            catch { /* ignorem */ }

            return null;
        }

        // ✅ En un UDO de document només cal cridar el NOM del fill tal com surt a OUDO/UDO1
        private SAPbobsCOM.GeneralDataCollection ResolveChildCollection(SAPbobsCOM.GeneralData header)
        {
            return header.Child("XNSEGCOM01"); // <-- fill: XNSEGCOM01
        }

        // ======================================================
        // 🔹 INSERTAR NUEVO SEGUIMIENTO
        // ======================================================
        public int Add(SegComHeader h, List<SegComLine> lines)
        {
            var svc = _company.GetCompanyService();
            var udo = (SAPbobsCOM.GeneralService)svc.GetGeneralService("XNSC");
            var head = (SAPbobsCOM.GeneralData)udo.GetDataInterface(
                SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

            // --- Capçalera ---
            head.SetProperty("U_CardCode", h.CardCode ?? "");
            head.SetProperty("U_CardName", h.CardName ?? "");
            head.SetProperty("U_NumAtCard", h.NumAtCard ?? "");
            head.SetProperty("U_DocDate", h.DocDate);                // datetime OK
            head.SetProperty("U_Status", (short)(h.Status));   // smallint → short
            head.SetProperty("U_BaseEntry", h.BaseEntry);  // int
                                                           //head.SetProperty("U_BaseNum", h.BaseNum);    // int
                                                           // U_BaseNum és nvarchar(50) → cal passar string
            if (h.BaseNum > 0)
                head.SetProperty("U_BaseNum", h.BaseNum.ToString());
            else
                head.SetProperty("U_BaseNum", ""); // o null si prefereixes no informar res

            // --- Línies ---
            var linesCol = ResolveChildCollection(head);               // "XNSEGCOM01"

            foreach (var l in lines)
            {
                var ln = linesCol.Add();
                ln.SetProperty("U_LineOrder", (short)l.LineOrder);    // smallint
                ln.SetProperty("U_Dscription", l.Dscription ?? "");

                // aquests camps són NVARCHAR a DB → passa-hi string
                ln.SetProperty("U_Date", (l.Date ?? DateTime.Now).ToString("yyyy-MM-dd"));
                ln.SetProperty("U_Hour", string.IsNullOrWhiteSpace(l.Hour)
                                                ? DateTime.Now.ToString("HH:mm")
                                                : l.Hour);
                ln.SetProperty("U_LineStatus", l.LineStatus ?? "");
                // Guardem StatusDate en el mateix format que el lector (yyyyMMdd HH:mm)
                ln.SetProperty("U_StatusDate", l.StatusDate?.ToString("yyyyMMdd HH:mm") ?? "");
            }

            try
{
    var res = udo.Add(head);
    return Convert.ToInt32(res.GetProperty("DocEntry"));
}
catch (Exception ex)
{
    var debugMsg = $"[ERROR ADD] {ex.Message}\n" +
                   $"CardCode={h.CardCode}, CardName={h.CardName}, " +
                   $"NumAtCard={h.NumAtCard}, DocDate={h.DocDate}, " +
                   $"Status={h.Status}, BaseEntry={h.BaseEntry}, BaseNum={h.BaseNum}";

    System.Diagnostics.Debug.WriteLine(debugMsg);
    throw new Exception($"SAP Error Add(): {ex.Message}");
}

        }

        public void Update(SegComHeader h, List<SegComLine> lines)
        {
            var svc = _company.GetCompanyService();
            var udo = (SAPbobsCOM.GeneralService)svc.GetGeneralService("XNSC");

            var p = (SAPbobsCOM.GeneralDataParams)udo.GetDataInterface(
                SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            p.SetProperty("DocEntry", h.DocEntry);

            var head = udo.GetByParams(p);

            head.SetProperty("U_NumAtCard", h.NumAtCard ?? "");
            head.SetProperty("U_Status", (short)(h.Status));

            var linesCol = ResolveChildCollection(head);               // "XNSEGCOM01"

            // Esborra totes les línies existents
            while (linesCol.Count > 0)
                linesCol.Remove(linesCol.Count - 1);

            // Torna a crear-les
            foreach (var l in lines)
            {
                var ln = linesCol.Add();
                ln.SetProperty("U_LineOrder", (short)l.LineOrder);
                ln.SetProperty("U_Dscription", l.Dscription ?? "");
                ln.SetProperty("U_Date", (l.Date ?? DateTime.Now).ToString("yyyy-MM-dd"));
                ln.SetProperty("U_Hour", string.IsNullOrWhiteSpace(l.Hour)
                                                ? DateTime.Now.ToString("HH:mm")
                                                : l.Hour);
                ln.SetProperty("U_LineStatus", l.LineStatus ?? "");
                //ln.SetProperty("U_StatusDate", l.StatusDate?.ToString("yyyy-MM-dd HH:mm") ?? "");
                ln.SetProperty("U_StatusDate", l.StatusDate?.ToString("yyyyMMdd HH:mm") ?? "");
            }

            udo.Update(head);
        }


        private DateTime GetServerDateTime()
        {
            var rs = (SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery("SELECT GETDATE() AS Now");
            return (DateTime)rs.Fields.Item("Now").Value;
        }

        public string GetServerTimestampString()
        {
            return GetServerDateTime().ToString("yyyyMMdd HH:mm");
        }

        // ======================================================
        // 🔹 OBTENER SEGUIMIENTO COMPLETO POR DOCENTRY (read)
        // ======================================================
        public SegComDocument GetByDocEntry(int docEntry)
        {
            var data = new SegComDocument();
            var query = new SegComQueryService(_company);

            var result = query.GetByDocEntry(docEntry);
            data.Header = result.Header;
            data.Lines = result.Lines;

            return data;
        }

        // ======================================================
        // 🔹 ELIMINAR SEGUIMIENTO (delete)
        // ======================================================
        public void Delete(int docEntry)
        {
            var svc = _company.GetCompanyService();
            var udo = (SAPbobsCOM.GeneralService)svc.GetGeneralService("XNSC");

            var p = (SAPbobsCOM.GeneralDataParams)udo.GetDataInterface(
                SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            p.SetProperty("DocEntry", docEntry);

            udo.Delete(p);
        }


    }
}


