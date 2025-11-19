using SAPbouiCOM; // UI API
using System;
using System.IO;
using System.Windows.Forms;
using XNSeguimentCompres.UI;

namespace XNSeguimentCompres
{
    internal static class Program
    {
        // ------------------------------------------------------------
        // Context SAP i instàncies principals
        // ------------------------------------------------------------
        private static SapContext ctx;
        private static SeguimentForm seguimentForm;

        // Alias curt per evitar escriure ctx.App contínuament
        private static SAPbouiCOM.Application SBO_Application => ctx.App;


        // ------------------------------------------------------------
        // Punt d’entrada. Ha de ser STA per la UI API (COM).
        // ------------------------------------------------------------
        [STAThread]
        private static void Main(string[] args)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

            ctx = new SapContext();

            try
            {
                // SAP llança el AddOn passant args[0] amb la connexió UI.
                if (args.Length == 0)
                {
#if DEBUG
                    MessageBox.Show(
                        "Aquest AddOn s'ha d'iniciar des de SAP Business One.",
                        "XNSeguimentCompres",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
#endif
                    return;
                }

                // Connexió UI
                ctx.ConnectViaGui(args[0]);

                // Tancar AddOn en sortir SAP
                ctx.App.AppEvent += (BoAppEventTypes EventType) =>
                {
                    if (EventType == BoAppEventTypes.aet_ShutDown ||
                        EventType == BoAppEventTypes.aet_CompanyChanged ||
                        EventType == BoAppEventTypes.aet_ServerTerminition)
                    {
                        try { System.Windows.Forms.Application.ExitThread(); } catch { }
                    }
                };

                // Missatge correcte en barra
                ctx.App.StatusBar.SetText(
                    "XNSeguimentCompres carregat correctament.",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success
                );

                // --------------------------------------------------------
                // Crear menú personalitzat
                // --------------------------------------------------------
                CrearMenu();
                SBO_Application.MenuEvent += Application_MenuEvent;

                // --------------------------------------------------------
                // Mantenir l’AddOn viu (loop UI)
                // --------------------------------------------------------
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error iniciant l'AddOn:\n\n" + ex.Message,
                    "XNSeguimentCompres",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }


        // ------------------------------------------------------------
        // Crear menú "XN Aplicacions" > "XN Seguiment compres"
        // ------------------------------------------------------------
        private static void CrearMenu()
        {
            try
            {
                var rootMenus = SBO_Application.Menus;
                var menuParams = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

                // Menú pare = "Mòduls" (43520)
                var modulesMenu = rootMenus.Item("43520");
                int position = modulesMenu.SubMenus.Count + 1;

                // Menú principal XN
                menuParams.Type = BoMenuType.mt_POPUP;
                menuParams.UniqueID = "XNApps";
                menuParams.String = "XN Aplicacions";
                menuParams.Position = position;

                string iconPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "imatges", "Logo_Quadrat_85Anys_16x16_OR-TR.png");
                if (File.Exists(iconPath))
                    menuParams.Image = iconPath;

                modulesMenu.SubMenus.AddEx(menuParams);

                // Submenú Seguiment compres
                var xnMenu = rootMenus.Item("XNApps");
                menuParams.Type = BoMenuType.mt_STRING;
                menuParams.UniqueID = "XNSegComp";
                menuParams.String = "XN Seguiment compres";
                menuParams.Position = 1;
                xnMenu.SubMenus.AddEx(menuParams);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(
                    "Error al crear menú: " + ex.Message,
                    BoMessageTime.bmt_Medium,
                    BoStatusBarMessageType.smt_Error
                );
            }
        }


        // ------------------------------------------------------------
        // Obrir formulari quan es selecciona el menú
        // ------------------------------------------------------------
        private static void Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.BeforeAction) return;
            if (pVal.MenuUID != "XNSegComp") return;

            try
            {
                if (seguimentForm == null)
                    seguimentForm = new SeguimentForm(ctx);

                string formPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "formularis",
                    "frmXNSeguimentCompres_new.xml"
                );

                if (File.Exists(formPath))
                {
                    seguimentForm.Load(formPath);
                }
                else
                {
                    SBO_Application.StatusBar.SetText(
                        "No s'ha trobat el formulari XML.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Error
                    );
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(
                    "Error obrint formulari: " + ex.Message,
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error
                );
            }

        }
    }
}

