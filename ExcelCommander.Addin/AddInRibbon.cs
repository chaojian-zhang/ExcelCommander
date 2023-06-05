using ExcelCommander.Base.ClientServer;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelCommander.Addin
{
    public partial class AddInRibbon
    {
        private void AddInRibbon_Load(object sender, RibbonUIEventArgs e){}

        #region Properties
        public int ServicePort { get; private set; }
        public Server Server { get; private set; }
        #endregion

        #region Service Control
        private void startButton_Click(object sender, RibbonControlEventArgs e)
        {
            Server = new Server(data => {
                CommandHandler handler = new CommandHandler();
                return handler.Handle(data);
            });
            ServicePort = Server.Start();
            statusLabel.Label = $"Service active on: {ServicePort}";

            startButton.Enabled = false;
            stopButton.Enabled = true;
        }

        private void stopButton_Click(object sender, RibbonControlEventArgs e)
        {
            Server.Stop();
            statusLabel.Label = "Service stopped.";

            startButton.Enabled = true;
            stopButton.Enabled = false;
        }
        #endregion
    }
}
