using NLog;

namespace ExplorerTabSwitcher
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            logger.Info("Started.");
            ApplicationConfiguration.Initialize();
            var dummyForm = new DummyForm(logger);
            Application.Run(dummyForm);
            logger.Info("Terminated.");
        }

        static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
    }
}