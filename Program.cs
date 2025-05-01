namespace EinhornExportIndex
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //License.LicenseKey = "IRONSUITE.ANDY.LUSHY.COM.10317-7C3597D93C-AGCOGQHN35O7TFV7-YW2YN3OXMLPX-ZVBLPJIYGHK6-ESCXUJFUJPFY-GAIWGYQENIDW-SNSGAJ6ZRKZQ-O6WDHG-TUIK2FDB5ZSOUA-DEPLOYMENT.TRIAL-CA6W7R.TRIAL.EXPIRES.22.MAR.2025";

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}