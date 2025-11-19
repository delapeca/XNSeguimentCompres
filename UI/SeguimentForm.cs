using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using XNSeguimentCompres.Application;
using XNSeguimentCompres.Data;
using XNSeguimentCompres.Domain;
using SAPButton = SAPbouiCOM.Button;
using SAPComboBox = SAPbouiCOM.ComboBox;
using SAPForm = SAPbouiCOM.Form;
using System.Globalization;
using ValidValues = SAPbouiCOM.ValidValues;
using System.Net.Http.Headers;

namespace XNSeguimentCompres.UI
{
    /// <summary>
    /// 🔹 Controla el formulari de Seguiment de Compres (UI)
    /// 🔹 Gestiona:
    ///     - Datasources de capçalera i línies
    ///     - Modes del formulari (ADD, OK, UPDATE, VER)
    ///     - Interacció amb serveis de negoci i repositoris
    ///     - Esdeveniments d’usuari i SAP
    /// </summary>
    public class SeguimentForm
    {
        // 🔹 Context SAP: Application + Company
        private readonly SapContext _ctx;

        // 🧠 Serveis de negoci i accés a dades
        private readonly SegComValidator _validator;
        private readonly SegComNumberingService _numbering;
        private readonly SegComRepository _repository;
        private readonly SegComQueryService _query;
        private readonly SegComApplicationService _appSvc;
        private SegComDocumentLoader _loader;

        // 🔒 Per no registrar events múltiples (idempotència)
        private bool _eventsSubscribed = false;

        // 🔀 Gestor d’estats del formulari
        private FormModeManager _mode;

        // 🔗 Objectes UI (còpies locals)
        private SAPForm _form;
        private DBDataSource _dsHead;
        private DBDataSource _dsLines;
        private Matrix _mtx;
        private SAPButton _btOk;
        private SAPButton _btCancel;
        private SAPbouiCOM.ButtonCombo _cbOk;

        /// <summary>
        /// Constructor
        /// Inicialitza serveis de negoci, NO carrega el formulari encara.
        /// </summary>
        public SeguimentForm(SapContext ctx)
        {
            _ctx = ctx;
            _validator = new SegComValidator();
            _numbering = new SegComNumberingService(ctx.Company);
            _repository = new SegComRepository(ctx.Company);
            _query = new SegComQueryService(ctx.Company);
            _appSvc = new SegComApplicationService(_validator, _repository);
        }

        /// <summary>
        /// 📌 Carrega i inicialitza el formulari Seguiment de Compres des d’un XML
        /// - Assigna datasources
        /// - Bindings controls
        /// - Crea i associa CFLs
        /// - Configura modes
        /// - Prepara matriu inicial
        /// - Registra esdeveniments
        /// </summary>
        public void Load(string xmlPath)
        {
            // 🧹 Tancar si hi ha una instància ja oberta
            try { _ctx.App.Forms.Item("XNSegComp").Close(); } catch { }

            // 📥 Obrir nova instància del formulari
            var xml = System.IO.File.ReadAllText(xmlPath);
            _ctx.App.LoadBatchActions(xml);
            _form = _ctx.App.Forms.Item("XNSegComp");
            _form.Visible = true;

            if (_form == null)
            {
                _ctx.App.StatusBar.SetText(
                    "[ERROR] _form és null després de carregar-lo!",
                    BoMessageTime.bmt_Long,
                    BoStatusBarMessageType.smt_Error);
                return;
            }

            _form.Freeze(true);
            try
            {
                // 📌 1️⃣ Connectar DataSources dels UDT
                _dsHead = _form.DataSources.DBDataSources.Add("@XNSEGCOM");
                _dsLines = _form.DataSources.DBDataSources.Add("@XNSEGCOM01");

                // 🔗 User DataSources per DocEntry/DocNum visual
                _form.DataSources.UserDataSources.Add("DE", BoDataType.dt_SHORT_TEXT, 20);
                _form.DataSources.UserDataSources.Add("DN", BoDataType.dt_SHORT_TEXT, 20);

                // 📌 2️⃣ Bindings dels controls
                ((EditText)_form.Items.Item("tDocEntry").Specific)
                    .DataBind.SetBound(true, "", "DE");
                ((EditText)_form.Items.Item("tDocNum").Specific)
                    .DataBind.SetBound(true, "", "DN");
                ((EditText)_form.Items.Item("tCardCode").Specific)
                    .DataBind.SetBound(true, "@XNSEGCOM", "U_CardCode");
                ((EditText)_form.Items.Item("tCardName").Specific)
                    .DataBind.SetBound(true, "@XNSEGCOM", "U_CardName");
                ((EditText)_form.Items.Item("tNumAtCard").Specific)
                    .DataBind.SetBound(true, "@XNSEGCOM", "U_NumAtCard");
                ((EditText)_form.Items.Item("tDocDate").Specific)
                    .DataBind.SetBound(true, "@XNSEGCOM", "U_DocDate");
                ((EditText)_form.Items.Item("tEstat").Specific)
                    .DataBind.SetBound(true, "@XNSEGCOM", "U_Status");

                // 📌 3️⃣ CFG ChooseFromList individuals
                CreateCfl("CFL_BPC", "22"); // BP per CardCode
                CreateCfl("CFL_BPN", "22"); // BP per Nom
                CreateCfl("CFL_OPOR", "22"); // Comandes compra
                CreateCfl("CFL_XNSC", "XNSC"); // Documents seguiment

                // 📌 4️⃣ Assigneu CFL als camps corresponents
                var edCardCode = (EditText)_form.Items.Item("tCardCode").Specific;
                edCardCode.ChooseFromListUID = "CFL_BPC";
                edCardCode.ChooseFromListAlias = "CardCode";

                var edCardName = (EditText)_form.Items.Item("tCardName").Specific;
                edCardName.ChooseFromListUID = "CFL_BPN";
                edCardName.ChooseFromListAlias = "CardName";

                var edDocNum = (EditText)_form.Items.Item("tDocNum").Specific;
                edDocNum.ChooseFromListUID = "CFL_OPOR";
                edDocNum.ChooseFromListAlias = "DocNum";

                var edDocEntry = (EditText)_form.Items.Item("tDocEntry").Specific;
                edDocEntry.ChooseFromListUID = "CFL_XNSC";
                edDocEntry.ChooseFromListAlias = "DocEntry";

                // 📌 5️⃣ Aplicar filtre: només OPOR obertes
                ApplyBpConditions("CFL_BPC");
                ApplyBpConditions("CFL_BPN");

                // 📌 6️⃣ Controls principals
                _mtx = (Matrix)_form.Items.Item("mtxLinies").Specific;
                _btOk = (SAPButton)_form.Items.Item("1").Specific;
                _btCancel = (SAPButton)_form.Items.Item("btCancel").Specific;
                _cbOk = (SAPbouiCOM.ButtonCombo)_form.Items.Item("cbOkAct").Specific;

                BindMatrixColumns();
                SetupOkCombo();

                // 📌 7️⃣ Mode inicial: ADD → propi formulari
                _mode = new FormModeManager(_form, _dsHead, _dsLines, _mtx, _cbOk, _btCancel, _numbering);
                _mode.SetNuevo();
                UpdateOkUiByMode();

                // 📌 8️⃣ Matriu: inicialitzar mínim 1 línia
                _mtx.FlushToDataSource();
                if (_dsLines.Size == 0)
                {
                    _dsLines.InsertRecord(0);
                    _dsLines.SetValue("U_LineOrder", 0, "1");
                }
                _mtx.LoadFromDataSource();

                // 📌 9️⃣ Combo d’estat (omple cStatus)
                ConfigureStatusColumn();

                // 📌 🔔 10️⃣ Subscriure esdeveniments (només 1 vegada)
                if (!_eventsSubscribed)
                {
                    _ctx.App.ItemEvent += App_ItemEvent;
                    _ctx.App.MenuEvent += App_MenuEvent;
                    _ctx.App.FormDataEvent += App_FormDataEvent;
                    _eventsSubscribed = true;
                }

                // 📌 11️⃣ Loader per carregar documents existents
                _loader = new SegComDocumentLoader(_query, _form, _dsHead, _dsLines, _mtx);

                // Bloc combo OK ha d'ignorar mode
                _form.Items.Item("cbOkAct").AffectsFormMode = false;
            }
            finally
            {
                _form.Freeze(false);
            }
        }

        /// <summary>
        /// Aplica les condicions del CFL de Proveïdors perquè només mostri
        /// aquells que tinguin comandes de compra OPOR obertes.
        /// </summary>
        /// <param name="cflId">ID del ChooseFromList a configurar</param>
        private void ApplyBpConditions(string cflId)
        {
            var cfl = (SAPbouiCOM.ChooseFromList)_form.ChooseFromLists.Item(cflId);
            var conds = (Conditions)_ctx.App.CreateObject(BoCreatableObjectType.cot_Conditions);

            // 🔹 Mostrem només documents amb estat Obert (DocStatus = 'O')
            var c = conds.Add();
            c.Alias = "DocStatus";
            c.Operation = BoConditionOperation.co_EQUAL;
            c.CondVal = "O";

            cfl.SetConditions(conds);
        }

        /// <summary>
        /// Neteja tots els valors vàlids d’un combo o columna abans de tornar-los a omplir.
        /// Evita valors duplicats o obsolets en la UI.
        /// </summary>
        private void ClearValidValues(ValidValues vals)
        {
            while (vals.Count > 0)
                vals.Remove(0, BoSearchKey.psk_Index);
        }

        /// <summary>
        /// Crea un ChooseFromList bàsic amb un UID, tipus d’objecte i multi-selecció opcional.
        /// S’afegeix al formulari actual.
        /// </summary>
        /// <param name="uid">Identificador únic del CFL</param>
        /// <param name="objectType">Tipus d’objecte SAP (22 = BP, etc.)</param>
        /// <param name="multi">Si permet multi-selecció</param>
        private void CreateCfl(string uid, string objectType, bool multi = false)
        {
            var cflParams = (ChooseFromListCreationParams)_ctx.App.CreateObject(
                BoCreatableObjectType.cot_ChooseFromListCreationParams);

            cflParams.UniqueID = uid;
            cflParams.ObjectType = objectType;
            cflParams.MultiSelection = multi;

            _form.ChooseFromLists.Add(cflParams);
        }

        /// <summary>
        /// Inicialitza el botó combo d’accions d’OK (Afegir i nou, veure, tancar).
        /// </summary>
        private void SetupOkCombo()
        {
            // 🔹 Neteja d’opcions prèvies
            ClearValidValues(_cbOk.ValidValues);

            // 🔹 Opcions disponibles després de guardar
            _cbOk.ValidValues.Add("ADDNEW", "Afegir i nou");
            _cbOk.ValidValues.Add("ADDVIEW", "Afegir i veure");
            _cbOk.ValidValues.Add("ADDCLOSE", "Afegir i tancar");

            // 🔹 Valor per defecte
            _cbOk.Select("ADDNEW", BoSearchKey.psk_ByValue);
            _cbOk.Caption = _cbOk.Selected.Description;
        }

        /// <summary>
        /// Controla la visibilitat/posició del botó OK estàndard i del combo personalitzat
        /// segons el mode del formulari:
        /// - ADD_MODE: es mostra el combo personalitzat
        /// - Altres modes: es mostra el botó OK/Actualitzar estàndard de SAP
        /// </summary>
        private void UpdateOkUiByMode()
        {
            var itOk = _form.Items.Item("1");
            var itCombo = _form.Items.Item("cbOkAct");

            if (_form.Mode == BoFormMode.fm_ADD_MODE)
            {
                // 🔁 Reaprofitem posició i mida del botó SAP per al combo
                itCombo.Left = itOk.Left;
                itCombo.Top = itOk.Top;
                itCombo.Width = itOk.Width;
                itCombo.Height = itOk.Height;

                itOk.Visible = false;
                itCombo.Visible = true;
            }
            else
            {
                itOk.Visible = true;
                itCombo.Visible = false;
            }
        }

        /// <summary>
        /// Configura la columna d’estat de línia (cStatus) de la matriu:
        /// - Intenta omplir valors des del UDT @XNSEGCOMLINESTATUS
        /// - Si falla, aplica valors per defecte (Pendent, Finalitzat, En espera)
        /// </summary>
        private void ConfigureStatusColumn()
        {
            var col = _mtx.Columns.Item("cStatus");
            ClearValidValues(col.ValidValues);

            try
            {
                var rs = (Recordset)_ctx.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(@"SELECT Code, Name FROM ""@XNSEGCOMLINESTATUS"" ORDER BY Code");

                while (!rs.EoF)
                {
                    col.ValidValues.Add(
                        rs.Fields.Item("Code").Value.ToString(),
                        rs.Fields.Item("Name").Value.ToString());
                    rs.MoveNext();
                }

                col.DisplayDesc = true;
            }
            catch
            {
                // 🔁 Valors per defecte si el UDT no existeix o dóna error
                col.ValidValues.Add("0", "Pendent");
                col.ValidValues.Add("1", "Finalitzat");
                col.ValidValues.Add("2", "En espera");
                col.DisplayDesc = true;
            }
        }

        /// <summary>
        /// Gestor central d’esdeveniments d’ítems:
        /// - Botons (OK, Cancel·lar, combo)
        /// - ChooseFromList (BP, OPOR, seguiments)
        /// - Validacions i modificacions de la matriu
        /// - Canvi de mode a UPDATE quan cal
        /// </summary>
        /// 
        private void App_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (_form == null || FormUID != _form.UniqueID)
                return;

            // ─────────────────────────────────────────────────────────────
            // 1️⃣ UPDATE_MODE (botó OK estàndard)
            // ─────────────────────────────────────────────────────────────
            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                pVal.ItemUID == "1" &&
                pVal.Before_Action &&
                _form.Mode == BoFormMode.fm_UPDATE_MODE)
            {
                HandleUpdateAction();
                BubbleEvent = false;
                return;
            }

            // ─────────────────────────────────────────────────────────────
            // 2️⃣ COMBO_SELECT → NOMÉS DEFINICIÓ D’ACCIÓ
            // ─────────────────────────────────────────────────────────────
            if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "cbOkAct")
            {
                if (pVal.Before_Action)
                {
                    Logger.Log("IGNORED COMBO_SELECT → Before_Action=true");
                    return;
                }

                // 🔹 Mantenim només el registre de la selecció
                string val = _cbOk?.Selected?.Value ?? "(null)";
                string desc = _cbOk?.Selected?.Description ?? "(null)";
                Logger.Log($"USER COMBO_SELECT → {val} : {desc}");

                _cbOk.Caption = desc;

                return;
            }

            // ─────────────────────────────────────────────────────────────
            // 3️⃣ EXECUCIÓ DE GUARDAR → ITEM_PRESSED en boto combo
            // ─────────────────────────────────────────────────────────────
            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                pVal.ItemUID == "cbOkAct" &&
                !pVal.Before_Action)
            {
                if (_cbOk.Selected == null)
                {
                    Logger.Log("ERROR: cbOkAct.ItemPressed sense Selected → Assigno ADDNEW");
                    _cbOk.Select("ADDNEW", BoSearchKey.psk_ByValue);
                }

                Logger.Log($"ITEM_PRESSED cbOkAct → Executant HandleOkAction({_cbOk.Selected.Value})");
                HandleOkAction();
                BubbleEvent = false;
                return;
            }

            // ─────────────────────────────────────────────────────────────
            // 4️⃣ BOTÓ CANCEL·LAR personalitzat
            // ─────────────────────────────────────────────────────────────
            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                pVal.ItemUID == "btCancel" &&
                !pVal.Before_Action)
            {
                HandleCancelAction();
                BubbleEvent = false;
                return;
            }

            // ─────────────────────────────────────────────────────────────
            // 5️⃣ CHOOSE FROM LIST
            // ─────────────────────────────────────────────────────────────
            if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.Before_Action)
            {
                var cflEv = (IChooseFromListEvent)pVal;
                if (cflEv.SelectedObjects == null) return;
                var row = cflEv.SelectedObjects;

                // 🔹 5.1 BP
                if (pVal.ItemUID == "tCardCode" || pVal.ItemUID == "tCardName")
                {
                    string bpCode = row.GetValue("CardCode", 0).ToString().Trim();
                    string bpName = row.GetValue("CardName", 0).ToString().Trim();
                    string docEntry = row.GetValue("DocEntry", 0).ToString().Trim();
                    string docNum = row.GetValue("DocNum", 0).ToString().Trim();
                    string numAtCard = row.GetValue("NumAtCard", 0).ToString().Trim();
                    string docStatus = row.GetValue("DocStatus", 0).ToString().Trim();
                    DateTime docDateDt = Convert.ToDateTime(row.GetValue("DocDate", 0));

                    string headerStatus = docStatus == "O" ? "0" :
                                          docStatus == "C" ? "1" : "2";

                    _form.Freeze(true);
                    try
                    {
                        _dsHead.SetValue("U_CardCode", 0, bpCode);
                        _dsHead.SetValue("U_CardName", 0, bpName);
                        _dsHead.SetValue("U_BaseEntry", 0, docEntry);
                        _dsHead.SetValue("U_BaseNum", 0, docNum);
                        _dsHead.SetValue("U_NumAtCard", 0, numAtCard);
                        _dsHead.SetValue("U_Status", 0, headerStatus);
                    }
                    finally
                    {
                        _form.Freeze(false);
                    }

                    if (int.TryParse(docEntry, out int baseEntryInt))
                    {
                        int existing = _query.GetSegComByBaseEntry(baseEntryInt);
                        if (existing > 0)
                        {
                            _form.Freeze(true);
                            try
                            {
                                _loader.Load(existing);
                                _mode.SetVer();
                                UpdateOkUiByMode();
                                _form.DataSources.UserDataSources.Item("DE").Value = existing.ToString();
                            }
                            finally
                            {
                                _form.Freeze(false);
                            }
                            return;
                        }
                    }

                    UpdateOkUiByMode();
                    return;
                }

                // 🔹 5.2 tDocEntry
                if (pVal.ItemUID == "tDocEntry")
                {
                    string docEntry = row.GetValue("DocEntry", 0).ToString().Trim();
                    _form.Freeze(true);
                    try
                    {
                        _form.DataSources.UserDataSources.Item("DE").Value = docEntry;
                        _loader.Load(int.Parse(docEntry));
                        _mode.SetVer();
                        UpdateOkUiByMode();
                    }
                    finally
                    {
                        _form.Freeze(false);
                    }
                    return;
                }

                // 🔹 5.3 tDocNum
                if (pVal.ItemUID == "tDocNum")
                {
                    string docEntry = row.GetValue("DocEntry", 0).ToString().Trim();
                    string docNum = row.GetValue("DocNum", 0).ToString().Trim();
                    string refNum = row.GetValue("NumAtCard", 0).ToString().Trim();
                    DateTime ddt = Convert.ToDateTime(row.GetValue("DocDate", 0));

                    _form.Freeze(true);
                    try
                    {
                        _dsHead.SetValue("U_BaseEntry", 0, docEntry);
                        _dsHead.SetValue("U_BaseNum", 0, docNum);
                        _dsHead.SetValue("U_NumAtCard", 0, refNum);
                    }
                    finally
                    {
                        _form.Freeze(false);
                    }

                    if (int.TryParse(docEntry, out int be2))
                    {
                        int existing = _query.GetSegComByBaseEntry(be2);
                        if (existing > 0)
                        {
                            _form.Freeze(true);
                            try
                            {
                                _loader.Load(existing);
                                _mode.SetVer();
                                UpdateOkUiByMode();
                                _form.DataSources.UserDataSources.Item("DE").Value = existing.ToString();
                            }
                            finally
                            {
                                _form.Freeze(false);
                            }
                            return;
                        }
                    }

                    UpdateOkUiByMode();
                    return;
                }
            }

            // ─────────────────────────────────────────────────────────────
            // 6️⃣ VALIDACIÓ COLUMNA DESCRIPCIÓ
            // ─────────────────────────────────────────────────────────────
            if (pVal.EventType == BoEventTypes.et_VALIDATE &&
                pVal.ItemUID == "mtxLinies" &&
                pVal.ColUID == "cDesc" &&
                !pVal.Before_Action)
            {
                _form.Freeze(true);
                try
                {
                    _mtx.FlushToDataSource();
                    int ds = pVal.Row - 1;
                    string desc = _dsLines.GetValue("U_Dscription", ds).Trim();

                    if (!string.IsNullOrWhiteSpace(desc))
                    {
                        _dsLines.SetValue("U_Date", ds, DateTime.Now.ToString("yyyy-MM-dd"));
                        _dsLines.SetValue("U_Hour", ds, DateTime.Now.ToString("HH:mm"));
                    }

                    int last = _dsLines.Size - 1;
                    if (ds == last && !string.IsNullOrWhiteSpace(desc))
                    {
                        _dsLines.InsertRecord(last + 1);
                        _dsLines.SetValue("U_LineOrder", last + 1, (last + 2).ToString());
                    }

                    _mtx.LoadFromDataSource();
                }
                finally
                {
                    _form.Freeze(false);
                }
                return;
            }

            // ─────────────────────────────────────────────────────────────
            // 7️⃣ VALIDACIÓ COLUMNA ESTAT
            // ─────────────────────────────────────────────────────────────
            if (pVal.EventType == BoEventTypes.et_VALIDATE &&
                pVal.ItemUID == "mtxLinies" &&
                pVal.ColUID == "cStatus" &&
                !pVal.Before_Action)
            {
                _mtx.FlushToDataSource();
                int ds = pVal.Row - 1;
                string desc = _dsLines.GetValue("U_Dscription", ds).Trim();
                string st = _dsLines.GetValue("U_LineStatus", ds).Trim();

                if (!string.IsNullOrWhiteSpace(desc) && string.IsNullOrWhiteSpace(st))
                {
                    BubbleEvent = false;
                    _ctx.App.StatusBar.SetText(
                        "Cal seleccionar un estat.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Error);
                }
                return;
            }
        }


        // Bloc substituit pel de sobre
        //private void App_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;

        //    // ✔️ Ignorem events d’altres formularis
        //    if (_form == null || FormUID != _form.UniqueID) return;

        //    // ─────────────────────────────────────────────────────────────
        //    // 1️⃣ Interceptar el botó OK estàndard en UPDATE_MODE (UPDATE)
        //    //     → Fem l’UPDATE manual via GeneralService i cancelem SAP
        //    // ─────────────────────────────────────────────────────────────
        //    if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
        //        pVal.ItemUID == "1" &&
        //        pVal.Before_Action &&
        //        _form.Mode == BoFormMode.fm_UPDATE_MODE)
        //    {
        //        HandleUpdateAction();

        //        // ❌ Evitem que SAP faci el seu UPDATE per defecte
        //        BubbleEvent = false;
        //        return;
        //    }

        //    // ─────────────────────────────────────────────────────────────
        //    // 2️⃣ Botó combo OK personalitzat (cbOkAct)
        //    //     → executa HandleOkAction només en acció d'usuari
        //    // ─────────────────────────────────────────────────────────────
        //    if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "cbOkAct")
        //    {
        //        // 🛡 Evitar COMBO_SELECT automàtics (refresh Mode / UI)
        //        if (pVal.Before_Action)
        //        {
        //            Logger.Log("IGNORED COMBO_SELECT → Before_Action = true (actualització SAP)");
        //            return;
        //        }

        //        // 🛡 Protecció completa contra nulls de SAP UI API
        //        try
        //        {
        //            // 🔹 Forçar SAP a actualitzar la selecció abans de llegir-la
        //            System.Threading.Thread.Sleep(50);

        //            // 🔹 Revalidar després del micro-wait
        //            if (_cbOk == null || _cbOk.Selected == null)
        //            {
        //                Logger.Log("ERROR: Després d'actualitzar encara Selected=null → assignant ADDNEW");
        //                _cbOk.Select("ADDNEW", BoSearchKey.psk_ByValue);
        //            }

        //            // 🔹 Guardem els valors un cop estables
        //            string optValue = _cbOk.Selected.Value;
        //            string optDesc = _cbOk.Selected.Description;

        //            Logger.Log($"USER COMBO_SELECT → {optValue} : {optDesc}");
        //            _cbOk.Caption = optDesc;

        //            Logger.Log("COMBO_SELECT → només registre, no guardem encara");
        //            return;


        //            //// 🚀 Ara sí → Executar l’acció seleccionada
        //            //Logger.Log("COMBO_SELECT → Executant HandleOkAction()");
        //            //HandleOkAction();
        //            //BubbleEvent = false;
        //            //return;


        //            // Bloc substituit pel codi de dalt
        //            //if (_cbOk == null || _cbOk.Selected == null)
        //            //{
        //            //    Logger.Log("IGNORED COMBO_SELECT → cbOk null o sense selecció");
        //            //    return;
        //            //}
        //            //
        //            //string selectedValue = _cbOk.Selected.Value ?? "(null)";
        //            //string selectedDesc = _cbOk.Selected.Description ?? "(null)";
        //            //Logger.Log($"USER COMBO_SELECT → {selectedValue} : {selectedDesc}");

        //            //_cbOk.Caption = selectedDesc;
        //        }
        //        catch (Exception ex)
        //        {
        //            // 🚑 Si aquí peta, SAPUI està canviant l'objecte al vol → NO continuem
        //            Logger.Log($"ERROR COMBO_SELECT ACCESS → {ex.Message}");
        //            return;
        //        }

        //        // 🔥 Si hem arribat aquí → és una acció vàlida d'usuari
        //        HandleOkAction();
        //        BubbleEvent = false;
        //        return;
        //    }

        //    // ─────────────────────────────────────────────────────────────
        //    // 3️⃣ Botó Cancel·lar personalitzat (btCancel)
        //    //     → Gestiona sortida, amb possible guardat previ
        //    // ─────────────────────────────────────────────────────────────
        //    if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
        //        pVal.ItemUID == "btCancel" &&
        //        !pVal.Before_Action)
        //    {
        //        HandleCancelAction();
        //        BubbleEvent = false;
        //        return;
        //    }

        //    // ─────────────────────────────────────────────────────────────
        //    // 4️⃣ Tractament dels ChooseFromList (CFL)
        //    //     - BP per codi o nom
        //    //     - Seguiment existent per DocEntry
        //    //     - Comanda OPOR per DocNum
        //    // ─────────────────────────────────────────────────────────────
        //    if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.Before_Action)
        //    {
        //        var cflEv = (IChooseFromListEvent)pVal;
        //        if (cflEv.SelectedObjects == null) return;
        //        var row = cflEv.SelectedObjects;

        //        // 🔹 4.1 Selecció de BP (tCardCode / tCardName)
        //        if (pVal.ItemUID == "tCardCode" || pVal.ItemUID == "tCardName")
        //        {
        //            string bpCode = row.GetValue("CardCode", 0).ToString().Trim();
        //            string bpName = row.GetValue("CardName", 0).ToString().Trim();
        //            string docNum = row.GetValue("DocNum", 0).ToString().Trim();
        //            string docEntry = row.GetValue("DocEntry", 0).ToString().Trim();

        //            DateTime docDateDt = Convert.ToDateTime(row.GetValue("DocDate", 0));
        //            string docDateDb = docDateDt.ToString("yyyyMMdd");
        //            string docDateUi = docDateDt.ToString("yyyy-MM-dd"); // (no necessari si el binding ja s’encarrega)

        //            string numAtCard = row.GetValue("NumAtCard", 0).ToString().Trim();
        //            string docStatus = row.GetValue("DocStatus", 0).ToString().Trim();

        //            // 🔁 Mapatge estat SAP → estat intern UDT (0/1/2)
        //            string headerStatus = docStatus == "O" ? "0" :
        //                                  docStatus == "C" ? "1" : "2";

        //            _form.Freeze(true);
        //            try
        //            {
        //                // 🔹 Capçalera @XNSEGCOM
        //                _dsHead.SetValue("U_CardCode", 0, bpCode);
        //                _dsHead.SetValue("U_CardName", 0, bpName);
        //                _dsHead.SetValue("U_BaseEntry", 0, docEntry);
        //                _dsHead.SetValue("U_BaseNum", 0, docNum);
        //                _dsHead.SetValue("U_NumAtCard", 0, numAtCard);
        //                _dsHead.SetValue("U_DocDate", 0, docDateDb);
        //                _dsHead.SetValue("U_Status", 0, docStatus);
        //                _dsHead.SetValue("U_Status", 0, headerStatus); // sobreescriu amb codi intern

        //                _form.DataSources.UserDataSources.Item("DN").Value = docNum;
        //                _form.Items.Item("mtxLinies").Update();

        //                // 🔹 Actualitzar condicions d'altres CFL segons BP
        //                ApplyPOConditions(bpCode);
        //                ApplySegComConditions(bpCode);
        //            }
        //            finally
        //            {
        //                _form.Freeze(false);
        //            }



        //            //// 🔐 NO carreguem seguiment existent automàticament en ADD_MODE
        //            //if (_form.Mode == BoFormMode.fm_ADD_MODE)
        //            //    return;

        //            // 🔍 En altres modes, si hi ha seguiment existent per aquesta OPOR, carregar-lo
        //            if (int.TryParse(docEntry, out int baseEntryInt))
        //            {
        //                int existingSegCom = _query.GetSegComByBaseEntry(baseEntryInt);

        //                if (existingSegCom > 0)
        //                {
        //                    _form.Freeze(true);
        //                    try
        //                    {
        //                        _loader.Load(existingSegCom);
        //                        _mode.SetVer();
        //                        UpdateOkUiByMode();
        //                        _form.DataSources.UserDataSources.Item("DE").Value = existingSegCom.ToString();
        //                    }
        //                    finally
        //                    {
        //                        _form.Freeze(false);
        //                    }

        //                    return;
        //                }
        //            }

        //            // Si no hi ha seguiment associat → assegurar que la UI està en mode correcte
        //            UpdateOkUiByMode();

        //            // 🔁 Convertir modificacions en UPDATE_MODE només si és un document existent
        //            if (pVal.EventType == BoEventTypes.et_VALIDATE &&
        //                !pVal.BeforeAction &&
        //                _form.Mode == BoFormMode.fm_OK_MODE &&
        //                !string.IsNullOrWhiteSpace(_form.DataSources.UserDataSources.Item("DE").Value))
        //            {
        //                _form.Mode = BoFormMode.fm_UPDATE_MODE;
        //                UpdateOkUiByMode();
        //            }

        //            return;
        //        }

        //        // 🔹 4.2 Selecció d’un seguiment existent (per DocEntry)
        //        if (pVal.ItemUID == "tDocEntry")
        //        {
        //            string docEntry = row.GetValue("DocEntry", 0).ToString().Trim();

        //            _form.Freeze(true);
        //            try
        //            {
        //                _form.DataSources.UserDataSources.Item("DE").Value = docEntry;
        //                _loader.Load(int.Parse(docEntry));
        //                _mode.SetVer();
        //                UpdateOkUiByMode();
        //            }
        //            finally
        //            {
        //                _form.Freeze(false);
        //            }

        //            return;
        //        }

        //        // 🔹 4.3 Selecció d’una comanda OPOR (per DocNum)
        //        if (pVal.ItemUID == "tDocNum")
        //        {
        //            _ctx.App.StatusBar.SetText(
        //                $"DEBUG: ABANS d’assignar OPOR → Mode={_form.Mode}",
        //                BoMessageTime.bmt_Short,
        //                BoStatusBarMessageType.smt_Warning);

        //            string docEntry = row.GetValue("DocEntry", 0).ToString().Trim();
        //            string docNum = row.GetValue("DocNum", 0).ToString().Trim();
        //            string docDate = Convert.ToDateTime(row.GetValue("DocDate", 0)).ToString("yyyy-MM-dd");
        //            string refNum = row.GetValue("NumAtCard", 0).ToString().Trim();

        //            _form.Freeze(true);
        //            _ctx.App.StatusBar.SetText(
        //                $"DEBUG: DESPRÉS d’assignar OPOR → Mode={_form.Mode}",
        //                BoMessageTime.bmt_Short,
        //                BoStatusBarMessageType.smt_Warning);

        //            try
        //            {
        //                _dsHead.SetValue("U_BaseEntry", 0, docEntry);
        //                _dsHead.SetValue("U_BaseNum", 0, docNum);
        //                _dsHead.SetValue("U_DocDate", 0, docDate);
        //                _dsHead.SetValue("U_NumAtCard", 0, refNum);
        //                _form.DataSources.UserDataSources.Item("DN").Value = docNum;
        //            }
        //            finally
        //            {
        //                _form.Freeze(false);
        //            }

        //            // 🔍 Comprovar si ja existeix un seguiment per aquesta base (OPOR)
        //            if (int.TryParse(docEntry, out int baseEntryInt))
        //            {
        //                int existingSegCom = _query.GetSegComByBaseEntry(baseEntryInt);

        //                if (existingSegCom > 0)
        //                {
        //                    _form.Freeze(true);
        //                    try
        //                    {
        //                        _ctx.App.StatusBar.SetText(
        //                            $"Ja existeix un seguiment per aquesta comanda. Carregant #{existingSegCom}...",
        //                            BoMessageTime.bmt_Short,
        //                            BoStatusBarMessageType.smt_Warning);

        //                        _loader.Load(existingSegCom);
        //                        _mode.SetVer();
        //                        UpdateOkUiByMode();
        //                        _form.DataSources.UserDataSources.Item("DE").Value = existingSegCom.ToString();
        //                    }
        //                    finally
        //                    {
        //                        _form.Freeze(false);
        //                    }

        //                    return;
        //                }
        //            }

        //            // 🟡 Si no existeix seguiment → mantenir Mode ADD (comboVisible)
        //            UpdateOkUiByMode();
        //            return;
        //        }
        //    }

        //    // ─────────────────────────────────────────────────────────────
        //    // 5️⃣ Matriu: validació de descripció de línia
        //    //     - Autocompletar data i hora
        //    //     - Crear nova línia quan cal
        //    // ─────────────────────────────────────────────────────────────
        //    if (pVal.EventType == BoEventTypes.et_VALIDATE &&
        //        pVal.ItemUID == "mtxLinies" &&
        //        pVal.ColUID == "cDesc" &&
        //        !pVal.Before_Action)
        //    {
        //        _form.Freeze(true);
        //        try
        //        {
        //            _mtx.FlushToDataSource();

        //            int rowUI = pVal.Row;
        //            int rowDS = rowUI - 1;
        //            string desc = _dsLines.GetValue("U_Dscription", rowDS).Trim();

        //            if (!string.IsNullOrWhiteSpace(desc))
        //            {
        //                _dsLines.SetValue("U_Date", rowDS, DateTime.Now.ToString("yyyy-MM-dd"));
        //                _dsLines.SetValue("U_Hour", rowDS, DateTime.Now.ToString("HH:mm"));
        //            }

        //            int last = _dsLines.Size - 1;
        //            if (!string.IsNullOrWhiteSpace(desc) && rowDS == last)
        //            {
        //                _dsLines.InsertRecord(last + 1);
        //                _dsLines.SetValue("U_LineOrder", last + 1, (last + 2).ToString());
        //            }

        //            _mtx.LoadFromDataSource();
        //            ((SAPComboBox)_mtx.Columns.Item("cStatus").Cells.Item(rowUI).Specific).Active = true;
        //        }
        //        finally
        //        {
        //            _form.Freeze(false);
        //        }

        //        return;
        //    }

        //    // ─────────────────────────────────────────────────────────────
        //    // 6️⃣ Matriu: validació d’estat per línia
        //    //     - Només valida si hi ha descripció
        //    //     - Obliga usuari a informar l’estat
        //    // ─────────────────────────────────────────────────────────────
        //    if (pVal.EventType == BoEventTypes.et_VALIDATE &&
        //        pVal.ItemUID == "mtxLinies" &&
        //        pVal.ColUID == "cStatus" &&
        //        !pVal.Before_Action)
        //    {
        //        _mtx.FlushToDataSource();
        //        int rowDS = pVal.Row - 1;

        //        string desc = _dsLines.GetValue("U_Dscription", rowDS).Trim();
        //        string status = _dsLines.GetValue("U_LineStatus", rowDS).Trim();

        //        if (!string.IsNullOrWhiteSpace(desc) &&
        //            string.IsNullOrWhiteSpace(status))
        //        {
        //            BubbleEvent = false;

        //            _ctx.App.StatusBar.SetText(
        //                "Cal seleccionar un estat.",
        //                BoMessageTime.bmt_Short,
        //                BoStatusBarMessageType.smt_Error);

        //            _mtx.Columns.Item("cStatus").Cells.Item(pVal.Row)
        //                .Click(BoCellClickType.ct_Regular);
        //        }

        //        return;
        //    }

        //    // ─────────────────────────────────────────────────────────────
        //    // 7️⃣ Canvis a la matriu en mode "veure" → passar a UPDATE_MODE
        //    // ─────────────────────────────────────────────────────────────
        //    if ((pVal.EventType == BoEventTypes.et_VALIDATE ||
        //         pVal.EventType == BoEventTypes.et_COMBO_SELECT) &&
        //        pVal.ItemUID == "mtxLinies" &&
        //        !pVal.Before_Action &&
        //        _form.Mode == BoFormMode.fm_OK_MODE)
        //    {
        //        _form.Mode = BoFormMode.fm_UPDATE_MODE;
        //        UpdateOkUiByMode();
        //    }

        //    // ─────────────────────────────────────────────────────────────
        //    // 8️⃣ Selecció d’estat de línia → assignar data d’estat i
        //    //     crear nova línia automàticament si cal.
        //    // ─────────────────────────────────────────────────────────────
        //    if (pVal.EventType == BoEventTypes.et_COMBO_SELECT &&
        //        pVal.ItemUID == "mtxLinies" &&
        //        pVal.ColUID == "cStatus" &&
        //        !pVal.Before_Action)
        //    {
        //        int curRowUI = pVal.Row;
        //        int curRowDS = curRowUI - 1;
        //        int nextRowUI = curRowUI + 1;

        //        _form.Freeze(true);
        //        try
        //        {
        //            _mtx.FlushToDataSource();

        //            DateTime now = DateTime.Now;

        //            // Guarda en format tècnic per BD
        //            _dsLines.SetValue("U_StatusDate", curRowDS, now.ToString("yyyyMMdd HH:mm"));

        //            int last = _dsLines.Size - 1;
        //            if (curRowDS == last)
        //            {
        //                _dsLines.InsertRecord(last + 1);
        //                _dsLines.SetValue("U_LineOrder", last + 1, (last + 2).ToString());

        //                // Estat per defecte = pendent
        //                _dsLines.SetValue("U_LineStatus", last + 1, "0");

        //                nextRowUI = curRowUI + 1;
        //            }

        //            _mtx.LoadFromDataSource();

        //            // Conversió a format amigable per UI (dd/MM/yyyy HH:mm)
        //            ((EditText)_mtx.Columns.Item("cStatDate")
        //                .Cells.Item(curRowUI).Specific).Value =
        //                now.ToString("dd/MM/yyyy HH:mm");
        //        }
        //        finally
        //        {
        //            _form.Freeze(false);
        //        }

        //        if (nextRowUI <= _mtx.RowCount)
        //        {
        //            ((EditText)_mtx.Columns.Item("cDesc")
        //                .Cells.Item(nextRowUI).Specific).Active = true;
        //        }

        //        return;
        //    }
        //}

        /// <summary>
        /// Gestió de menús estàndard SAP (ADD, FIND, navegació).
        /// Ajusta el mode i la UI dels botons segons l’acció.
        /// </summary>
        private void App_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (_form == null || pVal.BeforeAction) return;

            // 🔹 Menús que poden afectar el mode del formulari
            if (pVal.MenuUID == "1282" || // ADD
                pVal.MenuUID == "1281" || // FIND
                pVal.MenuUID == "1288" || // Prev
                pVal.MenuUID == "1289" || // Next
                pVal.MenuUID == "1290" || // First
                pVal.MenuUID == "1291")   // Last
            {
                _ctx.App.StatusBar.SetText(
                    $"[DEBUG MENU] Acció {pVal.MenuUID} → Mode resultant = {_form.Mode}",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Warning);
            }

            // 🆕 ADD MODE (nou seguiment)
            if (pVal.MenuUID == "1282")
            {
                _mode.SetNuevo();
                UpdateOkUiByMode();
                return;
            }

            // 🔁 FIND i navegació → només refresquem UI de botons
            if (pVal.MenuUID == "1281" || pVal.MenuUID == "1288" ||
                pVal.MenuUID == "1289" || pVal.MenuUID == "1290" ||
                pVal.MenuUID == "1291")
            {
                UpdateOkUiByMode();
                return;
            }
        }

        /// <summary>
        /// Vincula les columnes de la matriu als camps del DBDataSource de línies.
        /// </summary>
        private void BindMatrixColumns()
        {
            try
            {
                _mtx.Columns.Item("cOrder").DataBind.SetBound(true, "@XNSEGCOM01", "U_LineOrder");
                _mtx.Columns.Item("cDate").DataBind.SetBound(true, "@XNSEGCOM01", "U_Date");
                _mtx.Columns.Item("cHour").DataBind.SetBound(true, "@XNSEGCOM01", "U_Hour");
                _mtx.Columns.Item("cDesc").DataBind.SetBound(true, "@XNSEGCOM01", "U_Dscription");
                _mtx.Columns.Item("cStatDate").DataBind.SetBound(true, "@XNSEGCOM01", "U_StatusDate");
                _mtx.Columns.Item("cStatus").DataBind.SetBound(true, "@XNSEGCOM01", "U_LineStatus");
            }
            catch (Exception)
            {
                // En cas d’error de binding, es deixa silenciós per evitar trencar el formulari.
            }
        }

        /// <summary>
        /// Event de dades de formulari:
        /// - Quan SAP carrega dades d’un document (FORM_DATA_LOAD)
        ///   ajusta mode i UI a "veure" i amaga/mostra botons correctes.
        /// </summary>
        private void App_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (_form == null)
                return;

            if (BusinessObjectInfo.FormUID != _form.UniqueID)
                return;

            if (!BusinessObjectInfo.BeforeAction &&
                BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
            {
                var mode = _form.Mode;

                // 🔁 Ajusta la UI dels botons segons el mode actual
                UpdateOkUiByMode();

                // 📖 Si el document es troba en OK_MODE → el passem a mode "veure"
                if (mode == BoFormMode.fm_OK_MODE)
                {
                    _mode.SetVer();
                }
            }
        }

        /// <summary>
        /// Gestiona el guardat d’un nou seguiment en Mode ADD:
        /// - Construeix la capçalera (SegComHeader)
        /// - Prepara línies a partir del DBDataSource
        /// - Crida TryAdd() del servei d’aplicació
        /// - Aplica l’acció seleccionada al combo (Afegir i nou / veure / tancar)
        /// </summary>
        /// 
        private void HandleOkAction()
        {
            try
            {
                // 🛑 DESACTIVAR EVENTS per evitar bucle ADD
                Logger.Log("ItemEvent UNSUBSCRIBED → inici Add action");
                _ctx.App.ItemEvent -= App_ItemEvent;

                // 🔒 Bloquejar botons mentre es desa
                _btCancel.Item.Enabled = false;
                _cbOk.Item.Enabled = false;

                // 🔹 Construir capçalera per ADD
                var header = new SegComHeader
                {
                    DocEntry = int.TryParse(_dsHead.GetValue("DocEntry", 0), out var de) ? de : 0,
                    DocNum = int.TryParse(_dsHead.GetValue("U_BaseNum", 0), out var dn) ? dn : 0,
                    CardCode = _dsHead.GetValue("U_CardCode", 0).Trim(),
                    CardName = _dsHead.GetValue("U_CardName", 0).Trim(),
                    NumAtCard = _dsHead.GetValue("U_NumAtCard", 0).Trim(),
                    Status = int.TryParse(_dsHead.GetValue("U_Status", 0).Trim(), out var s) ? s : 0,
                    BaseEntry = int.TryParse(_dsHead.GetValue("U_BaseEntry", 0), out var be) ? be : 0,
                    BaseNum = int.TryParse(_dsHead.GetValue("U_BaseNum", 0), out var bn) ? bn : 0,
                    DocDate = DateTime.Today
                };

                // 🔥 FIX: assegurar que sempre hi hagi una línia buida al final
                _mtx.FlushToDataSource();

                int last = _dsLines.Size - 1;
                bool lastHasData =
                    !string.IsNullOrWhiteSpace(_dsLines.GetValue("U_Dscription", last)) ||
                    !string.IsNullOrWhiteSpace(_dsLines.GetValue("U_LineStatus", last));

                if (lastHasData)
                {
                    _dsLines.InsertRecord(last + 1);
                    _dsLines.SetValue("U_LineOrder", last + 1, (last + 2).ToString());
                }

                _mtx.LoadFromDataSource();
                _mtx.FlushToDataSource();

                // 🔨 Construir línies
                var lines = new List<SegComLine>();

                for (int i = 0; i < _dsLines.Size; i++)
                {
                    var desc = _dsLines.GetValue("U_Dscription", i).Trim();
                    if (string.IsNullOrEmpty(desc)) continue;

                    lines.Add(new SegComLine
                    {
                        LineOrder = int.Parse(_dsLines.GetValue("U_LineOrder", i)),
                        Date = DateTime.TryParse(_dsLines.GetValue("U_Date", i), out var date)
                                    ? date : (DateTime?)null,
                        Hour = _dsLines.GetValue("U_Hour", i).Trim(),
                        Dscription = desc,
                        LineStatus = _dsLines.GetValue("U_LineStatus", i).Trim(),
                        StatusDate = DateTime.TryParseExact(
                                        _dsLines.GetValue("U_StatusDate", i).Trim(),
                                        "dd/MM/yyyy HH:mm",
                                        CultureInfo.InvariantCulture,
                                        DateTimeStyles.None,
                                        out var sd)
                                    ? sd : (DateTime?)null
                    });
                }

                // 🧠 Si ja hi ha DocEntry → això NO és un ADD, és un UPDATE
                if (header.DocEntry > 0)
                {
                    Logger.Log("DETECTED UPDATE MODE → delegant a HandleUpdateAction()");
                    HandleUpdateAction();
                    return;
                }

                // 🔹 Save ADD
                if (!_appSvc.TryAdd(header, lines, out int newDocEntry, out var err))
                {
                    _ctx.App.StatusBar.SetText(err,
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Warning);
                    return;
                }

                if (_cbOk == null || _cbOk.Selected == null)
                {
                    _ctx.App.StatusBar.SetText(
                        "Advertència: el botó d'acció no tenia cap opció seleccionada. S'ha assignat 'Afegir i nou'.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Warning);

                    _cbOk.Select("ADDNEW", BoSearchKey.psk_ByValue);
                }

                // 🔀 Accions combo
                switch (_cbOk.Selected.Value)
                {
                    case "ADDNEW":
                        _mode.SetNuevo();
                        UpdateOkUiByMode();
                        break;

                    case "ADDVIEW":
                        _form.DataSources.UserDataSources.Item("DE").Value = newDocEntry.ToString();
                        _loader.Load(newDocEntry);
                        ((EditText)_form.Items.Item("tDocEntry").Specific).Value = newDocEntry.ToString();
                        _mode.SetVer();
                        UpdateOkUiByMode();
                        break;

                    case "ADDCLOSE":
                        _form.Close();
                        break;
                }

                _ctx.App.StatusBar.SetText(
                    "Document desat correctament.",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                _ctx.App.StatusBar.SetText(
                    "[ERROR] Guardant document: " + ex.Message,
                    BoMessageTime.bmt_Long,
                    BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                // 🔓 Reactivar botons després del procés
                _btCancel.Item.Enabled = true;
                _cbOk.Item.Enabled = true;

                // 🔄 REACTIVAR EVENTS un cop acabada la lògica
                _ctx.App.ItemEvent += App_ItemEvent;
                Logger.Log("ItemEvent RESUBSCRIBED → fi Add action");
            }
        }

        /// <summary>
        /// Gestiona l’actualització (UPDATE) d’un seguiment existent:
        /// - Llegeix la capçalera i línies actuals
        /// - Construeix entitats de domini
        /// - Crida TryUpdate() del servei d’aplicació
        /// - Torna el formulari a mode "veure"
        /// </summary>
        private void HandleUpdateAction()
        {
            try
            {
                // 🔒 Bloquejar botons mentre actualitzem
                _btCancel.Item.Enabled = false;
                _btOk.Item.Enabled = false;

                // 1️⃣ Construir capçalera per UPDATE
                int docEntry = int.TryParse(_dsHead.GetValue("DocEntry", 0), out var de) ? de : 0;
                if (docEntry <= 0)
                {
                    _ctx.App.StatusBar.SetText(
                        "[ERROR] No s'ha pogut determinar el DocEntry del seguiment a actualitzar.",
                        BoMessageTime.bmt_Long,
                        BoStatusBarMessageType.smt_Error);
                    return;
                }

                var header = new SegComHeader
                {
                    DocEntry = docEntry,
                    CardCode = _dsHead.GetValue("U_CardCode", 0).Trim(),
                    CardName = _dsHead.GetValue("U_CardName", 0).Trim(),
                    NumAtCard = _dsHead.GetValue("U_NumAtCard", 0).Trim(),
                    Status = int.TryParse(_dsHead.GetValue("U_Status", 0).Trim(), out var s) ? s : 0,
                    BaseEntry = int.TryParse(_dsHead.GetValue("U_BaseEntry", 0), out var be) ? be : 0,
                    BaseNum = int.TryParse(_dsHead.GetValue("U_BaseNum", 0), out var bn) ? bn : 0,
                    DocDate = DateTime.Today
                };

                // 2️⃣ Sincronitzar matriu → DataSource
                _mtx.FlushToDataSource();

                // ─────────────────────────────────────────────────────────
                // 🧱 Construir col·lecció de línies per UPDATE
                // ─────────────────────────────────────────────────────────
                // ─── DUPLICAT INICI (patró similar a HandleOkAction) ───
                var lines = new List<SegComLine>();

                for (int i = 0; i < _dsLines.Size; i++)
                {
                    var desc = _dsLines.GetValue("U_Dscription", i).Trim();
                    if (string.IsNullOrEmpty(desc))
                        continue;

                    lines.Add(new SegComLine
                    {
                        LineOrder = int.TryParse(_dsLines.GetValue("U_LineOrder", i), out var lo) ? lo : (i + 1),
                        Date = DateTime.TryParse(_dsLines.GetValue("U_Date", i), out var date)
                                    ? date
                                    : (DateTime?)null,
                        Hour = _dsLines.GetValue("U_Hour", i).Trim(),
                        Dscription = desc,
                        LineStatus = _dsLines.GetValue("U_LineStatus", i).Trim(),
                        StatusDate = DateTime.TryParseExact(
                                        _dsLines.GetValue("U_StatusDate", i).Trim(),
                                        "dd/MM/yyyy HH:mm",
                                        CultureInfo.InvariantCulture,
                                        DateTimeStyles.None,
                                        out var sd)
                                    ? sd
                                    : (DateTime?)null // DUPLICAT (parsing StatusDate també a HandleOkAction)
                    });
                }
                // ─── DUPLICAT FI ───

                // 3️⃣ Cridar servei d’aplicació per fer UPDATE
                if (!_appSvc.TryUpdate(header, lines, out var err))
                {
                    _ctx.App.StatusBar.SetText(
                        err,
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Warning);
                    return;
                }

                // 4️⃣ Tornar a mode "veure"
                _mode.SetVer();
                UpdateOkUiByMode();

                _ctx.App.StatusBar.SetText(
                    "Document actualitzat correctament.",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                _ctx.App.StatusBar.SetText(
                    "[ERROR] Actualitzant document: " + ex.Message,
                    BoMessageTime.bmt_Long,
                    BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                _btCancel.Item.Enabled = true;
                _btOk.Item.Enabled = true;
            }
        }

        /// <summary>
        /// Gestió del botó Cancel·lar:
        /// - Si hi ha canvis en ADD/UPDATE → pregunta si es vol desar
        /// - Pot desar i després tancar, tancar sense desar, o cancel·lar l’acció
        /// </summary>
        private void HandleCancelAction()
        {
            try
            {
                if (_form == null || !_form.Visible)
                    return;

                bool hasPendingChanges =
                    _form.Mode == BoFormMode.fm_ADD_MODE ||
                    _form.Mode == BoFormMode.fm_UPDATE_MODE;

                if (hasPendingChanges)
                {
                    int answer = _ctx.App.MessageBox(
                        "Tens canvis sense desar.\nVols guardar-los abans de tancar el formulari?",
                        3,
                        "Sí", "No", "Cancel·lar");

                    if (answer == 1)
                    {
                        // Sí: desar i tancar
                        HandleOkAction();
                        _form.Close();
                    }
                    else if (answer == 2)
                    {
                        // No: tancar sense desar
                        _form.Close();
                    }
                    else
                    {
                        // Cancel·lar: no fer res
                        return;
                    }
                }
                else
                {
                    // Cap canvi pendent → tancar directament
                    _form.Close();
                }
            }
            catch (Exception ex)
            {
                _ctx.App.StatusBar.SetText(
                    "[ERROR] Cancel·lar: " + ex.Message,
                    BoMessageTime.bmt_Long,
                    BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Aplica condicions al CFL de comandes de compra (OPOR):
        /// - Només comandes del proveïdor indicat
        /// - Només documents oberts (DocStatus = 'O')
        /// </summary>
        private void ApplyPOConditions(string cardCode)
        {
            var cfl = (SAPbouiCOM.ChooseFromList)_form.ChooseFromLists.Item("CFL_OPOR");
            var conds = (Conditions)_ctx.App.CreateObject(BoCreatableObjectType.cot_Conditions);

            var c = conds.Add();
            c.Alias = "CardCode";
            c.Operation = BoConditionOperation.co_EQUAL;
            c.CondVal = cardCode;

            c = conds.Add();
            c.Alias = "DocStatus";
            c.Operation = BoConditionOperation.co_EQUAL;
            c.CondVal = "O";

            cfl.SetConditions(conds);
        }

        /// <summary>
        /// Aplica condicions al CFL de seguiments existents (XNSC)
        /// filtrant per proveïdor (U_CardCode).
        /// </summary>
        private void ApplySegComConditions(string bpCode)
        {
            var cfl = (SAPbouiCOM.ChooseFromList)_form.ChooseFromLists.Item("CFL_XNSC");
            var conds = (Conditions)_ctx.App.CreateObject(BoCreatableObjectType.cot_Conditions);

            if (!string.IsNullOrWhiteSpace(bpCode))
            {
                var c = conds.Add();
                c.Alias = "U_CardCode";
                c.Operation = BoConditionOperation.co_EQUAL;
                c.CondVal = bpCode;
            }

            cfl.SetConditions(conds);
        }
    }
}

