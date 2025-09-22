// Program.cs
using System.ServiceProcess;
namespace ExportLimbusService
{
    static class Program
    {
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new LimbusExportService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}