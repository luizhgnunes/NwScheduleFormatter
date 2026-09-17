using Velopack;
using Velopack.Sources;
using System.Windows;

namespace NwScheduleFormatter.Services
{
    public static class UpdateService
    {
        public static async Task CheckForUpdatesAsync()
        {
            var mgr = new UpdateManager(new GithubSource("https://github.com/luizhgnunes/NwScheduleFormatter", null, false));
            UpdateInfo? newVersion = null;
            try
            {
                newVersion = await mgr.CheckForUpdatesAsync();
            }
            catch
            {
                // Swallow exceptions to avoid crashing the application if update check fails
            }
            if (newVersion == null) 
                return; // No updates available

            var result = MessageBox.Show("Nova versão disponível. Deseja baixar?", "Atualização", MessageBoxButton.YesNo);
            if (result == MessageBoxResult.Yes)
            {
                await mgr.DownloadUpdatesAsync(newVersion);
                mgr.ApplyUpdatesAndRestart(newVersion);
            }
        }
    }
}
