using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Reflection;

namespace CleverAge.OdfConverter.OdfWord2003Addin
{
    [RunInstaller(true)]
    public partial class DependencyChecker : Installer
    {
        public DependencyChecker() {
            InitializeComponent();
        }

        public override void Install(System.Collections.IDictionary stateSaver) {
            base.Install(stateSaver);
            bool cancelInstall = false;
            foreach (AssemblyName assembly in Assembly.GetExecutingAssembly().GetReferencedAssemblies()) {
                try {
                    Assembly ass = Assembly.Load(assembly);
                } catch  {
                    cancelInstall = true;
                }
            }
            if (cancelInstall) {
                FrmPrerequisites helperDialog = new FrmPrerequisites();
                helperDialog.ShowDialog();
                throw new InstallException("Installation will be cancelled");
            }
        }
    }
}