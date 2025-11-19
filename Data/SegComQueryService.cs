using System;
using System.Collections.Generic;
using SAPbobsCOM;
using XNSeguimentCompres.Domain;

namespace XNSeguimentCompres.Data
{
    /// <summary>
    /// Servicio de consultas (lectura) de seguimientos y pedidos.
    /// 
    /// * No guarda ni modifica datos. *
    /// * No contiene lógica de negocio. *
    /// * Reutilizable desde UI, API, o app móvil. *
    /// </summary>
    public class SegComQueryService
    {
        private readonly Company _company;

        public SegComQueryService(Company company)
        {
            _company = company;
        }

        // ======================================================
        // 🔹 OBTENER UN SEGUIMIENTO COMPLETO POR DOCENTRY
        // ======================================================
        public (SegComHeader Header, List<SegComLine> Lines) GetByDocEntry(int docEntry)
        {
            var header = GetHeader(docEntry);
            var lines = GetLinesByDocEntry(docEntry);
            return (header, lines);
        }

        // ======================================================
        // 🧾 CABECERA
        // ======================================================
        public SegComHeader GetHeader(int docEntry)
        {
            Recordset rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery($@"
                SELECT DocEntry, DocNum, U_CardCode, U_CardName, U_DocDate, 
                       U_NumAtCard, U_Status
                FROM ""@XNSEGCOM""
                WHERE DocEntry = {docEntry}");

            if (rs.RecordCount == 0)
                return null;

            return new SegComHeader
            {
                DocEntry = docEntry,
                DocNum = Convert.ToInt32(rs.Fields.Item("DocNum").Value),
                CardCode = rs.Fields.Item("U_CardCode").Value.ToString().Trim(),
                CardName = rs.Fields.Item("U_CardName").Value.ToString().Trim(),
                DocDate = Convert.ToDateTime(rs.Fields.Item("U_DocDate").Value),
                NumAtCard = rs.Fields.Item("U_NumAtCard").Value.ToString().Trim(),
                Status = (int)rs.Fields.Item("U_Status").Value
            };
        }

        /// <summary>
        /// Retorna el DocEntry del seguiment associat a una OPOR concreta.
        /// Si no existeix, retorna 0.
        /// </summary>
        public int GetSegComByBaseEntry(int baseEntry)
        {
            Recordset rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery($@"
                SELECT DocEntry 
                FROM ""@XNSEGCOM""
                WHERE U_BaseEntry = {baseEntry}");

            if (rs.RecordCount == 0)
                return 0;

            return Convert.ToInt32(rs.Fields.Item("DocEntry").Value);
        }

        // ======================================================
        // 📋 LÍNEAS DEL SEGUIMIENTO
        // ======================================================
        public List<SegComLine> GetLinesByDocEntry(int docEntry)
        {
            var list = new List<SegComLine>();

            Recordset rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery($@"SELECT  DocEntry,
                            LineId,
                            U_Dscription,
                            U_LineOrder,
                            U_Date,
                            U_Hour,
                            U_LineStatus,
                            U_StatusDate
                    FROM [@XNSEGCOM01]
                    WHERE DocEntry = {docEntry}
                    ORDER BY LineId");

            //rs.DoQuery($@"
            //    SELECT LineID, U_Dscription, U_Date, U_Hour, U_LineOrder, U_LineStatus
            //    FROM ""@XNSEGCOM01""
            //    WHERE DocEntry = {docEntry}
            //    ORDER BY U_LineOrder, LineID");

            while (!rs.EoF)
            {
                var vLineId = rs.Fields.Item("LineId").Value;
                var vDesc = rs.Fields.Item("U_Dscription").Value;
                var vDate = rs.Fields.Item("U_Date").Value;
                var vHour = rs.Fields.Item("U_Hour").Value;
                var vOrder = rs.Fields.Item("U_LineOrder").Value;
                var vStatus = rs.Fields.Item("U_LineStatus").Value;
                var vStatusDate = rs.Fields.Item("U_StatusDate").Value;

                list.Add(new SegComLine
                {
                    LineId = Convert.ToInt32(vLineId),

                    Dscription = vDesc == null || vDesc is DBNull ? "" : vDesc.ToString(),

                    Date = vDate == null || vDate is DBNull || vDate.ToString().Trim() == ""
                            ? (DateTime?)null
                            : Convert.ToDateTime(vDate),

                    Hour = vHour == null || vHour is DBNull ? "" : vHour.ToString(),

                    LineOrder = Convert.ToInt32(vOrder),

                    LineStatus = vStatus == null || vStatus is DBNull ? "" : vStatus.ToString(),

                    StatusDate =
                        vStatusDate == null || vStatusDate is DBNull || vStatusDate.ToString().Trim() == ""
                        ? (DateTime?)null
                        : DateTime.ParseExact(
                            vStatusDate.ToString(),
                            "yyyyMMdd HH:mm",
                            System.Globalization.CultureInfo.InvariantCulture
                          )
                });
                rs.MoveNext();
            }

            return list;
        }

        // ======================================================
        // 🔍 SEGUIMIENTOS POR PROVEEDOR (Opcional para UI Buscar)
        // ======================================================
        public List<SegComHeader> FindByCardCode(string cardCode)
        {
            var list = new List<SegComHeader>();
            Recordset rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery($@"
                SELECT DocEntry, U_DocNum, U_CardCode, U_CardName, U_DocDate, U_Status
                FROM ""@XNSEGCOM""
                WHERE U_CardCode = '{cardCode.Replace("'", "''")}'
                ORDER BY DocEntry DESC");

            while (!rs.EoF)
            {
                list.Add(new SegComHeader
                {
                    DocEntry = Convert.ToInt32(rs.Fields.Item("DocEntry").Value),
                    DocNum = Convert.ToInt32(rs.Fields.Item("U_DocNum").Value),
                    CardCode = cardCode,
                    CardName = rs.Fields.Item("U_CardName").Value.ToString(),
                    DocDate = Convert.ToDateTime(rs.Fields.Item("U_DocDate").Value),
                    Status = rs.Fields.Item("U_Status").Value.ToString()
                });

                rs.MoveNext();
            }

            return list;
        }

        // ======================================================
        // 📦 PEDIDOS DE COMPRA ABIERTOS (OPOR)
        // ======================================================
        public List<(int DocEntry, int DocNum, DateTime DocDate)> GetOpenPurchaseOrders(string cardCode)
        {
            var list = new List<(int, int, DateTime)>();
            Recordset rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

            rs.DoQuery($@"
                SELECT DocEntry, DocNum, DocDate
                FROM OPOR
                WHERE CardCode = '{cardCode.Replace("'", "''")}'
                AND DocStatus = 'O'
                ORDER BY DocDate DESC");

            while (!rs.EoF)
            {
                list.Add((
                    Convert.ToInt32(rs.Fields.Item("DocEntry").Value),
                    Convert.ToInt32(rs.Fields.Item("DocNum").Value),
                    Convert.ToDateTime(rs.Fields.Item("DocDate").Value)
                ));

                rs.MoveNext();
            }

            return list;
        }
    }
}
