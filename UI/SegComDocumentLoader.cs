using SAPbouiCOM;
using XNSeguimentCompres.Data;

namespace XNSeguimentCompres.UI
{
    /// <summary>
    /// Se encarga de cargar un documento desde la base de datos
    /// y reflejarlo visualmente en el formulario SAP.
    /// </summary>
    public class SegComDocumentLoader
    {
        private readonly SegComQueryService _query;
        private readonly Form _form;
        private readonly DBDataSource _dsHead;
        private readonly DBDataSource _dsLines;
        private readonly Matrix _mtx;

        public SegComDocumentLoader(
            SegComQueryService query,
            Form form,
            DBDataSource dsHead,
            DBDataSource dsLines,
            Matrix mtx)
        {
            _query = query;
            _form = form;
            _dsHead = dsHead;
            _dsLines = dsLines;
            _mtx = mtx;
        }

        public void Load(int docEntry)
        {
            _form.Freeze(true);
            try
            {
                var data = _query.GetByDocEntry(docEntry);

                // ──────────────────────────
                // CAPÇALERA
                // ──────────────────────────
                _dsHead.SetValue("DocEntry", 0, docEntry.ToString());
                _dsHead.SetValue("DocNum", 0, data.Header.DocNum.ToString());
                _dsHead.SetValue("U_CardCode", 0, data.Header.CardCode);
                _dsHead.SetValue("U_CardName", 0, data.Header.CardName);
                _dsHead.SetValue("U_DocDate", 0, data.Header.DocDate.ToString("yyyyMMdd"));
                _dsHead.SetValue("U_NumAtCard", 0, data.Header.NumAtCard ?? "");

                // ──────────────────────────
                // LÍNIES EXISTENTS
                // ──────────────────────────
                _dsLines.Clear();
                int i = 0;
                foreach (var l in data.Lines)
                {
                    if (i >= _dsLines.Size) _dsLines.InsertRecord(i);
                    _dsLines.SetValue("LineID", i, l.LineId.ToString());
                    _dsLines.SetValue("U_Dscription", i, l.Dscription);
                    _dsLines.SetValue("U_Date", i, l.Date?.ToString("yyyyMMdd") ?? "");
                    _dsLines.SetValue("U_Hour", i, l.Hour);
                    _dsLines.SetValue("U_LineOrder", i, l.LineOrder.ToString());
                    _dsLines.SetValue("U_LineStatus", i, l.LineStatus);

                    // 👁️ Format llegible per UI
                    _dsLines.SetValue("U_StatusDate", i,
                        l.StatusDate?.ToString("dd/MM/yyyy HH:mm") ?? "");

                    i++;
                }

                // ──────────────────────────
                // ➕ NOVA LÍNIA BUIDA automàtica
                // ──────────────────────────
                _dsLines.InsertRecord(i);
                _dsLines.SetValue("U_LineOrder", i, (i + 1).ToString()); // ordre consecutiu

                // ──────────────────────────
                // Actualitzar matriu UI i focus
                // ──────────────────────────
                _mtx.LoadFromDataSource();
                

                // Focus a descripció de la línia nova 🎯
                if (_mtx.RowCount > 0)
                {
                    ((SAPbouiCOM.EditText)_mtx.Columns.Item("cDesc")
                        .Cells.Item(_mtx.RowCount).Specific).Active = true;
                }

                // Per mostrar botons correctes
                _form.Mode = BoFormMode.fm_OK_MODE;
            }
            finally
            {
                _form.Freeze(false);
            }
        }
    }
}

