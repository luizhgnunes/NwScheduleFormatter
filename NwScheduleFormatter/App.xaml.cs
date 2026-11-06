using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using NwScheduleFormatter.Data;
using NwScheduleFormatter.Data.Interfaces;
using NwScheduleFormatter.Data.Repositories;
using NwScheduleFormatter.Data.Repositories.Interfaces;
using NwScheduleFormatter.Services;
using NwScheduleFormatter.Services.Interfaces;
using System.IO;
using System.Windows;
using Velopack;

namespace NwScheduleFormatter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static SuperintendentVisitSchedule SuperintendentVisitSchedule { get; private set; }
        public static IServiceProvider? ServiceProvider { get; private set; }

        protected override async void OnStartup(StartupEventArgs e)
        {
            VelopackApp.Build().Run();

            // Configure dependency injection
            var services = new ServiceCollection();
            ConfigureServices(services);
            ServiceProvider = services.BuildServiceProvider();

            // Initialize database
            InitializeDatabase();

            SuperintendentVisitSchedule = new SuperintendentVisitSchedule();
            var jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "superintendent-visits.json");
            SuperintendentVisitSchedule.LoadFromJson(jsonPath);

            var whatsappService = ServiceProvider.GetRequiredService<IWhatsappService>();
            await whatsappService.LoadWhatsappCredentials();

            // Create MainWindow via DI
            var mainWindow = ServiceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();

            base.OnStartup(e);
        }

        /// <summary>
        /// Configures the application services.
        /// </summary>
        private void ConfigureServices(IServiceCollection services)
        {
            var dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "App.db");
            var connectionString = $"Data Source={dbPath};Cache=Shared";

            services.AddDbContext<NwScheduleFormatterDbContext>(options =>
                options.UseSqlite(connectionString));

            services.AddScoped<MainWindow>();
            services.AddScoped<IJwMeetingReader, JwMeetingReader>();
            services.AddSingleton<IWhatsappService, WhatsappService>();
            services.AddSingleton<IWhatsAppConfigRepository, WhatsAppConfigRepository>();
            services.AddScoped<IWhatsAppMessagingService, WhatsAppMessagingService>();
            services.AddScoped<ISendingOfDesignationRepository, SendingOfDesignationRepository>();
            services.AddScoped<IDesignationContactRepository, DesignationContactRepository>();
            services.AddScoped<IWhatsAppConfigRepository, WhatsAppConfigRepository>();
            services.AddScoped<IPdfGenerationService, PdfGenerationService>();
            services.AddScoped<ICsvParsingService, CsvParsingService>();
        }

        /// <summary>
        /// Initializes the database, creating tables if necessary.
        /// </summary>
        private void InitializeDatabase()
        {
            try
            {
                using (var scope = ServiceProvider?.CreateScope())
                {
                    var context = scope?.ServiceProvider.GetRequiredService<NwScheduleFormatterDbContext>();
                    context?.Database.EnsureCreated();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing database: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

}

