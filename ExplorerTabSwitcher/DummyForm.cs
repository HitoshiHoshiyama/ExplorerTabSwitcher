using NLog;

namespace ExplorerTabSwitcher
{
    public partial class DummyForm : Form
    {
        public DummyForm(Logger logger)
        {
            InitializeComponent();

            this.logger = logger;
            this.Width = 200;
            this.Height = 150;
            this.ShowInTaskbar = false;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.AllowTransparency = true;
            this.FormClosing += OnClosing;

#if DEBUG
            this.Opacity = 0.5;
#else
            this.Opacity = 0;
#endif
            this.mouseHook = new MouseHook(this.logger);
            this.logger.Info("Create DummyForm.");
        }

        private void OnClosing(object? sender, FormClosingEventArgs e)
        {
            this.mouseHook.Dispose();
            this.logger.Info("Close DummyForm.");
        }

        private Logger logger;

        private MouseHook mouseHook;
    }
}
