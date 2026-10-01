using System.Windows;
using System.IO;
using Velopack;
using NwScheduleFormatter.Data;

namespace NwScheduleFormatter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static SuperintendentVisitSchedule SuperintendentVisitSchedule { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            VelopackApp.Build().Run();

            SuperintendentVisitSchedule = new SuperintendentVisitSchedule();
            var jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "superintendent-visits.json");
            SuperintendentVisitSchedule.LoadFromJson(jsonPath);

            base.OnStartup(e);
        }
    }

}
