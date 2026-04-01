using NwScheduleFormatter.Helpers;
using NwScheduleFormatter.Models;
using PuppeteerSharp;
using System.Text.RegularExpressions;

namespace NwScheduleFormatter.Data;

public class JwMeetingReader
{
    private const string JW_BASE_URL = "https://www.jw.org";
    private const string JW_WORKBOOK_URL_SUFFIX = "/pt/biblioteca/jw-apostila-do-mes/";

    private string _workbookListHtml;

    public async Task<JwWebsiteMeeting> ReadFromJwWebsiteAsync(DateOnly mondayDate)
    {
        var url = GetUrlForMonth(mondayDate);
        if (string.IsNullOrWhiteSpace(_workbookListHtml))
            _workbookListHtml = await GetHtmlAsync(url);

        var workbookUrlSuffix =  Regex
                .Match(_workbookListHtml, $@"<a href=""(\S+)"">\s*{DateHelper.GetWeekInFull(mondayDate)}").Groups[1].Value;

        var workbookHtml = await GetHtmlAsync($"{JW_BASE_URL}{workbookUrlSuffix}");

        var meeting = new JwWebsiteMeeting();
        meeting.InitialSong.Number = Convert.ToInt16(Regex.Match(workbookHtml, @">Cântico\s+(\d+)</strong></a> <strong>e oração \| Comentários iniciais").Groups[1].Value);
        meeting.MiddleSong.Number = Convert.ToInt16(Regex.Match(workbookHtml, @">Cântico\s+(\d+)</strong></a></h3>").Groups[1].Value);
        meeting.FinalSong.Number = Convert.ToInt16(Regex.Match(workbookHtml, @">Cântico\s+(\d+)</a></span> e oração</h3>").Groups[1].Value);
        meeting.Apply1DurationMinutes = Convert.ToInt16(Regex.Match(workbookHtml, @">4\.[\s\S]*?\((\d+) min\)").Groups[1].Value);
        meeting.Apply2DurationMinutes = Convert.ToInt16(Regex.Match(workbookHtml, @">5\.[\s\S]*?\((\d+) min\)").Groups[1].Value);
        meeting.Apply3DurationMinutes = Convert.ToInt16(Regex.Match(workbookHtml, @">6\.[\s\S]*?\((\d+) min\)").Groups[1].Value);

        var nextPartNumber = 7;
        var matchApply4 = Regex.Match(workbookHtml, @">7\.[\s\S]*?\((\d+) min\)[\s\S]*?NOSSA VIDA CRISTÃ</h2>");
        if (matchApply4.Success)
        {
            meeting.Apply4DurationMinutes = Convert.ToInt16(matchApply4.Groups[1].Value);
            nextPartNumber++;
        }

        meeting.LivingAsChristhians1DurationMinutes = Convert.ToInt16(Regex.Match(workbookHtml, $@">{nextPartNumber}\.[\s\S]*?\((\d+) min\)").Groups[1].Value);
        nextPartNumber++;
        if (meeting.LivingAsChristhians1DurationMinutes < 15)
            meeting.LivingAsChristhians2DurationMinutes = Convert.ToInt16(Regex.Match(workbookHtml, $@">{nextPartNumber}\.[\s\S]*?\((\d+) min\)").Groups[1].Value);

        return meeting;
    }

    private static string GetUrlForMonth(DateOnly mondayDate)
    {
        var urlSufix = string.Empty;
        switch (mondayDate.Month)
        {
            case 1:
            case 2:
                urlSufix = $"janeiro-fevereiro-{mondayDate.Year}-mwb";
                break;
            case 3:
            case 4:
                urlSufix = $"marco-abril-{mondayDate.Year}-mwb";
                break;
            case 5:
            case 6:
                urlSufix = $"maio-junho-{mondayDate.Year}-mwb";
                break;
            case 7:
            case 8:
                urlSufix = $"julho-agosto-{mondayDate.Year}-mwb";
                break;
            case 9:
            case 10:
                urlSufix = $"setembro-outubro-{mondayDate.Year}-mwb";
                break;
            case 11:
            case 12:
                urlSufix = $"novembro-dezembro-{mondayDate.Year}-mwb";
                break;
        }
        return $"{JW_BASE_URL}{JW_WORKBOOK_URL_SUFFIX}{urlSufix}";
    }

    private static async Task<string> GetHtmlAsync(string url)
    {
        // Configura o download do navegador (necessário apenas na primeira execução)
        var browserFetcher = new BrowserFetcher();
        await browserFetcher.DownloadAsync();

        // Lança o navegador em modo "headless" (escondido)
        var launchOptions = new LaunchOptions
        {
            Headless = true,
            Args = new[] { "--no-sandbox", "--disable-setuid-sandbox" }
        };

        try
        {
            using var browser = await Puppeteer.LaunchAsync(launchOptions);
            using var page = await browser.NewPageAsync();

            // Simula um User-Agent de um navegador real para evitar bloqueios
            await page.SetUserAgentAsync("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");

            // Tenta navegar até a URL
            var navigationOptions = new NavigationOptions
            {
                WaitUntil = new[] { WaitUntilNavigation.Networkidle2 }, // Aguarda o tráfego de rede parar
                Timeout = 60000 // 60 segundos
            };

            await page.GoToAsync(url, navigationOptions);

            // Opcional: Aguarda um seletor específico se a página demorar a renderizar
            // await page.WaitForSelectorAsync(".main-content"); 

            // Retorna o HTML completo após a renderização do JavaScript
            return await page.GetContentAsync();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao acessar a página: {ex.Message}");
            throw;
        }
    }
}
