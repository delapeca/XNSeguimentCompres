using SAPbouiCOM;
using XNSeguimentCompres.Domain;

namespace XNSeguimentCompres.UI
{
    /// <summary>
    /// Se encarga exclusivamente de preparar el formulario SAP en cada modo.
    /// No contiene lógica de negocio, solo configuración de UI.
    /// </summary>
    public class FormModeManager
    {
        private readonly Form _form;
        private readonly DBDataSource _dsHead;
        private readonly DBDataSource _dsLines;
        private readonly Matrix _mtx;
        private readonly ButtonCombo _cbOk;
        private readonly Button _btCancel;
        private readonly SegComNumberingService _numbering;


        public FormModeManager(
            Form form,
            DBDataSource dsHead,
            DBDataSource dsLines,
            Matrix mtx,
            ButtonCombo cbOk,
            Button btCancel,
            SegComNumberingService numbering)
        {
            _form = form;
            _dsHead = dsHead;
            _dsLines = dsLines;
            _mtx = mtx;
            _cbOk = cbOk;
            _btCancel = btCancel;
            _numbering = numbering;
        }

        /// <summary>
        /// Pone el formulario listo para crear un nuevo seguimiento.
        /// </summary>
        public void SetNuevo()
        {

            _form.Freeze(true);
            try
            {
                _form.Mode = BoFormMode.fm_ADD_MODE;

                _dsHead.Clear();
                _dsLines.Clear();

                if (_dsHead.Size == 0) _dsHead.InsertRecord(0);
                if (_dsLines.Size == 0) _dsLines.InsertRecord(0);

                _dsHead.SetValue("DocNum", 0, _numbering.GetNextDocNumFromUdt().ToString());
                ((SAPbouiCOM.EditText)_form.Items.Item("tDocEntry").Specific).Value = "";


                _dsLines.Clear();

                while (_dsLines.Size > 0)
                    _dsLines.RemoveRecord(0);

                _dsLines.InsertRecord(_dsLines.Size);
                _dsHead.SetValue("U_CardCode", 0, "");
                _dsHead.SetValue("U_CardName", 0, "");
                _dsLines.SetValue("U_LineOrder", _dsLines.Size - 1, "1");

                _mtx.LoadFromDataSource();

                // 🔥 Rehabilitar edició de columnes (després de Search mode)
                foreach (Column col in _mtx.Columns)
                {
                    // Les que han de ser NO editables
                    if (col.UniqueID == "cLineId" || col.UniqueID == "cStatDate")
                        col.Editable = false;
                    else
                        col.Editable = true; // ← aquesta és la clau que es perdia
                }

                // 👁️ Tornar a recarregar després de reactivar editabilitat
                _mtx.LoadFromDataSource();


                while (_cbOk.ValidValues.Count > 0)
                    _cbOk.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);

                _cbOk.ValidValues.Add("1", "Afegir i Nou");
                _cbOk.ValidValues.Add("2", "Afegir i Veure");
                _cbOk.ValidValues.Add("3", "Afegir i Tancar");
                _cbOk.Select("1", BoSearchKey.psk_ByValue);

                // Establir que només es mostri la descripció
                _cbOk.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                _form.Items.Item("1").Visible = false;
                _form.Items.Item("cbOkAct").Visible = true;

                _btCancel.Caption = "Cancel·lar";
                _btCancel.Item.Visible = true;


                //--------------------------------------------------------------
                // FIX: assegurar que el combo té un valor seleccionat en mode nou
                //--------------------------------------------------------------
                try
                {
                    _cbOk.Select("1", BoSearchKey.psk_ByValue);
                    _cbOk.Caption = _cbOk.Selected.Description;
                }
                catch
                {
                    // ignorem: només fallback
                }

            }
            finally
            {
                _form.Freeze(false);
            }
        }

        /// <summary>
        /// Muestra el documento cargado en modo lectura.
        /// </summary>
        public void SetVer()
        {
            _form.Freeze(true);
            try
            {
                _form.Mode = BoFormMode.fm_OK_MODE;

                _form.Items.Item("1").Visible = true;
                _form.Items.Item("cbOkAct").Visible = false;

                _btCancel.Item.Visible = true;
            }
            finally
            {
                _form.Freeze(false);
            }
        }
    }
}

