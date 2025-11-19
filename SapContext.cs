using SAPbobsCOM;
using SAPbouiCOM;
using System;


namespace XNSeguimentCompres
{
    public class SapContext
    {
        public SAPbouiCOM.Application App { get; private set; }
        public SAPbobsCOM.Company Company { get; private set; }

        public void ConnectViaGui(string sboGuiApiConnStr)
        {
            var guiApi = new SboGuiApi();
            guiApi.Connect(sboGuiApiConnStr);            // p.ej. viene de args[0]
            App = guiApi.GetApplication(-1);

            Company = new SAPbobsCOM.Company();
            Company = (SAPbobsCOM.Company)App.Company.GetDICompany(); // Reutiliza la sesión del cliente
            if (Company.Connected == false)
                throw new Exception("No se pudo obtener la conexión DI desde el cliente.");
        }

        // Alternativa: conexión directa (si no lanzas desde el cliente)
        public void ConnectDirect(string server, string db, string user, string pass, string slUser, string slPass, BoDataServerTypes dbType)
        {
            Company = new SAPbobsCOM.Company
            {
                Server = server,
                CompanyDB = db,
                DbServerType = dbType,
                UserName = user,
                Password = pass,
                language = BoSuppLangs.ln_Spanish,
                UseTrusted = false,
                LicenseServer = slUser, // o "srv:30000"
                // SL user/pass si usas SLD; si no, comenta estas líneas
            };
            var rc = Company.Connect();
            if (rc != 0) throw new Exception($"Error DI: {Company.GetLastErrorDescription()}");
        }

        public Recordset Rs() => (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);
    }
}
