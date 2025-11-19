using System;
using System.IO;

namespace XNSeguimentCompres.Application
{
    /// <summary>
    /// Logger simple a fitxer per diagnòstic del comportament del formulari.
    /// NO usa UI SAP → no penja l'aplicació si entra en bucle.
    /// </summary>
    public static class Logger
    {
        private static readonly object _lock = new object();

        /// <summary>
        /// Escriu una línia de log al fitxer local.
        /// Crea el directori si no existeix.
        /// </summary>
        public static void Log(string message)
        {
            try
            {
                string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string file = Path.Combine(folder, "XNSeguimentCompres.log");
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  {message}";

                lock (_lock)
                {
                    File.AppendAllText(file, line + Environment.NewLine);
                }
            }
            catch
            {
                // NI UNA EXCEPCIÓ: mai frenar execució
            }
        }
    }
}

