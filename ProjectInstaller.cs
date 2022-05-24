using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;

namespace SysFileWatcher_new
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        protected override void OnBeforeInstall(IDictionary savedState)
        {
            //string parameter = "N:\\Temp\\Bart\" \"FileWatcherLog";
           // Context.Parameters["assemblypath"] = "\"" + Context.Parameters["assemblypath"] + "\" \"" + parameter + "\"";
            base.OnBeforeInstall(savedState);
        }

        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void serviceInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {

        }
    }
}
