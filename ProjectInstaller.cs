// ProjectInstaller.cs
using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;
namespace ExportLimbusService
{
    [RunInstaller(true)]
    public class ProjectInstaller : Installer
    {
        private ServiceProcessInstaller _processInstaller;
        private ServiceInstaller _serviceInstaller;
        public ProjectInstaller()
        {
            _processInstaller = new ServiceProcessInstaller
            {
                Account = ServiceAccount.LocalSystem
            };
            _serviceInstaller = new ServiceInstaller
            {
                ServiceName = "LimbusExport",
                DisplayName = "LimbusExportService",
                Description = "Service d'export vers le réseau des fichiers RTStruct Limbus DICOM",
                StartType = ServiceStartMode.Automatic
            };
            Installers.Add(_processInstaller);
            Installers.Add(_serviceInstaller);
        }
    }
}