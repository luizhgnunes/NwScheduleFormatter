using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using NwScheduleFormatter.Configuration;
using NwScheduleFormatter.Helpers;
using NwScheduleFormatter.Models;
using NwScheduleFormatter.Services.Interfaces;
using PdfSharp.Fonts;
using System;
using System.IO;
using Colors = MigraDoc.DocumentObjectModel.Colors;
using Section = MigraDoc.DocumentObjectModel.Section;
using Style = MigraDoc.DocumentObjectModel.Style;
using Table = MigraDoc.DocumentObjectModel.Tables.Table;

namespace NwScheduleFormatter.Services;

/// <summary>
/// Service for generating PDF documents related to meeting schedules.
/// </summary>
public class PdfGenerationService : IPdfGenerationService
{
    private int _selectedYear;
    private int _selectedMonth;

    /// <summary>
    /// Sets the context for PDF generation (year and month).
    /// </summary>
    public void SetContext(int year, int month)
    {
        _selectedYear = year;
        _selectedMonth = month;
    }

    public async Task GenerateMeetingSchedulePdfAsync(string outputDirectory, List<Meeting> meetings, IProgress<int> progress, int progressSlice)
    {
        var superintendentVisitSchedule = App.SuperintendentVisitSchedule;

        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var document = new Document();
        document.Info.Title = "Programação da Reunião do Meio de Semana";
        document.Info.Author = ApplicationSettings.CONGREGATION_NAME;

        DefineStyles(document);

        Section section = document.AddSection();
        section.PageSetup.PageFormat = PageFormat.A4;
        section.PageSetup.LeftMargin = "2.0cm";
        section.PageSetup.RightMargin = "2.0cm";
        section.PageSetup.TopMargin = "0.8cm";
        section.PageSetup.BottomMargin = "0.8cm";

        for (var i = 0; i < meetings.Count; i++)
        {
            if (i > 1 && (i + 1) % 2 > 0)
                section.AddPageBreak();

            var addBlankRow = false;
            if ((i + 1) % 2 > 0)
                WritePageHeader(section);
            else
                addBlankRow = true;

            WriteDateAndPresidentRow(section, DateHelper.GetWeekInFull(meetings[i].Date).ToUpper(), meetings[i].Chairman, meetings[i].WeeklyBibleReading, addBlankRow);

            var table = section.AddTable();
            table.AddColumn(Unit.FromCentimeter(2));
            table.AddColumn(Unit.FromCentimeter(8));
            table.AddColumn(Unit.FromCentimeter(3));
            table.AddColumn(Unit.FromCentimeter(4));

            var timeSplit = ApplicationSettings.START_TIME.Split(':');
            var time = new TimeOnly(Convert.ToInt32(timeSplit[0]), Convert.ToInt32(timeSplit[1]));

            AddRow(table, time.ToString("HH:mm"), $"• Cântico {meetings[i].InitialSong.Number}", "Oração", meetings[i].OpeningPrayer);
            time = time.AddMinutes(5);
            AddRow(table, time.ToString("HH:mm"), "• Comentários iniciais (1 min)");

            WriteHeaderRow(table, "TESOUROS DA PALAVRA DE DEUS", Color.FromRgb(87, 90, 93));
            time = time.AddMinutes(1);
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].TreasuresFromGodsWord.Treasures.Theme} (10 min)", "", meetings[i].TreasuresFromGodsWord.Treasures.Speaker);
            time = time.AddMinutes(10);
            AddRow(table, time.ToString("HH:mm"), "2. Joias espirituais (10 min)", "", meetings[i].TreasuresFromGodsWord.SpiritualGems);
            time = time.AddMinutes(10);
            AddRow(table, time.ToString("HH:mm"), "3. Leitura da Bíblia (4 min)", "", meetings[i].TreasuresFromGodsWord.BibleReading);

            WriteHeaderRow(table, "FAÇA SEU MELHOR NO MINISTÉRIO", Color.FromRgb(190, 137, 0));
            
            time = time.AddMinutes(5);
            var apply1Assistant = string.IsNullOrWhiteSpace(meetings[i].Apply1.Assistant + 1) ? string.Empty : $" / {meetings[i].Apply1.Assistant}";
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].Apply1.Theme} ({meetings[i].Apply1.DurationMinutes} min)", "", $"{meetings[i].Apply1.Speaker}{apply1Assistant}");
            
            time = time.AddMinutes((int)meetings[i].Apply1.DurationMinutes + 1);
            var apply2Assistant = string.IsNullOrWhiteSpace(meetings[i].Apply2.Assistant) ? string.Empty : $" / {meetings[i].Apply2.Assistant}";
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].Apply2.Theme} ({meetings[i].Apply2.DurationMinutes} min)", "", $"{meetings[i].Apply2.Speaker}{apply2Assistant}");
            
            time = time.AddMinutes((int)meetings[i].Apply2.DurationMinutes + 1);
            var apply3Assistant = string.IsNullOrWhiteSpace(meetings[i].Apply3.Assistant) ? string.Empty : $" / {meetings[i].Apply3.Assistant}";
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].Apply3.Theme} ({meetings[i].Apply3.DurationMinutes} min)", "", $"{meetings[i].Apply3.Speaker}{apply3Assistant}");
            
            short partNumber = 7;
            if (!string.IsNullOrWhiteSpace(meetings[i].Apply4.Speaker))
            {
                time = time.AddMinutes((int)meetings[i].Apply3.DurationMinutes);
                var apply4Assistant = string.IsNullOrWhiteSpace(meetings[i].Apply4.Assistant) ? string.Empty : $" / {meetings[i].Apply4.Assistant}";
                AddRow(table, time.ToString("HH:mm"), $"{meetings[i].Apply4.Theme} ({meetings[i].Apply4.DurationMinutes} min)", "", $"{meetings[i].Apply4.Speaker}{apply4Assistant}", partNumber);
                partNumber++;
            }

            WriteHeaderRow(table, "NOSSA VIDA CRISTÃ", Color.FromRgb(126, 0, 36));
            int previousPartMinutes = string.IsNullOrWhiteSpace(meetings[i].Apply4.Speaker) ? (int)meetings[i].Apply3.DurationMinutes : (int)meetings[i].Apply4.DurationMinutes;

            time = time.AddMinutes(previousPartMinutes + 1);
            AddRow(table, time.ToString("HH:mm"), $"• Cântico {meetings[i].MiddleSong.Number}");
            time = time.AddMinutes(5);
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].LivingAsChristhians1.Theme} ({meetings[i].LivingAsChristhians1.DurationMinutes} min)", "", $"{meetings[i].LivingAsChristhians1.Speaker}", partNumber);
            partNumber++;
            if (!string.IsNullOrWhiteSpace(meetings[i].LivingAsChristhians2.Speaker))
            {
                time = time.AddMinutes((int)meetings[i].LivingAsChristhians1.DurationMinutes);
                AddRow(table, time.ToString("HH:mm"), $"{meetings[i].LivingAsChristhians2.Theme} ({meetings[i].LivingAsChristhians2.DurationMinutes} min)", "", $"{meetings[i].LivingAsChristhians2.Speaker}", partNumber);
                partNumber++;
            }

            previousPartMinutes = string.IsNullOrWhiteSpace(meetings[i].LivingAsChristhians2.Speaker) ? (int)meetings[i].LivingAsChristhians1.DurationMinutes : (int)meetings[i].LivingAsChristhians2.DurationMinutes;
            time = time.AddMinutes(previousPartMinutes);

            var isSuperintendentVisit = superintendentVisitSchedule.IsSuperintendentVisit(meetings[i].Date);

            if (!isSuperintendentVisit)
            {
                AddRow(table, time.ToString("HH:mm"), $"{partNumber}. Estudo bíblico de congregação (30 min)", "Dirigente/leitor", $"{meetings[i].CongregationBibleStudy.Speaker} / {meetings[i].CongregationBibleStudy.Reader}");
                time = time.AddMinutes(30);
            }
            
            AddRow(table, time.ToString("HH:mm"), "• Comentários finais (3 min)");
            time = time.AddMinutes(3);

            if (isSuperintendentVisit)
            {
                var superintendentVisit = superintendentVisitSchedule.GetSuperintendentVisitSpeech(meetings[i].Date);
                AddRow(table, time.ToString("HH:mm"), $"{partNumber}. {superintendentVisit.Theme} (30 min)", "Orador", $"{superintendentVisit.Speaker}");
                time = time.AddMinutes(30);
                if ((superintendentVisit.Song?.Number ?? 0) > 0)
                {
                    meetings[i].FinalSong.Number = superintendentVisit.Song.Number;
                    meetings[i].FinalSong.Description = superintendentVisit.Song.Description;
                }
            }

            AddRow(table, time.ToString("HH:mm"), $"• Cântico {meetings[i].FinalSong.Number}", "Oração", $"{meetings[i].ClosingPrayer}");

            var progressValue = (int)(((i + 1) / (double)meetings.Count) * progressSlice);
            progress?.Report(progressValue + 80);
        }
        var monthFileName = _selectedMonth.ToString().PadLeft(2, '0');
        var outputPath = Path.Combine(outputDirectory, $"{_selectedYear}-{monthFileName} S-140_T Programação Reunião Meio de Semana.pdf");
        FileService.SaveFile(document, outputPath);
        progress?.Report(80 + progressSlice);
    }

    public async Task GenerateStudentDesignationsPdfAsync(string outputDirectory, List<Meeting> meetings, IProgress<int> progress, int progressSlice)
    {
        var monthFileName = _selectedMonth.ToString().PadLeft(2, '0');
        var fileService = new FileService();
        var designationsOutputPath = Path.Combine(outputDirectory, $"{_selectedYear}-{monthFileName} Designações Reunião Meio de Semana.pdf");
        fileService.CreateStudentDesignationDocuments(meetings, designationsOutputPath, progress, progressSlice);
        progress?.Report(100);
    }

    private void WritePageHeader(Section section)
    {
        var table = section.AddTable();
        table.Borders.Bottom.Width = 1;
        table.AddColumn(Unit.FromCentimeter(4.7));
        table.AddColumn(Unit.FromCentimeter(12.3));

        var row = table.AddRow();
        row.HeightRule = RowHeightRule.AtLeast;
        row.Height = Unit.FromCentimeter(1);
        row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;

        var timeCell0 = row.Cells[0];
        timeCell0.AddParagraph("EURICO CHAVES");
        timeCell0.Format.Font = new Font("Arial", 11);
        timeCell0.Format.Font.Bold = true;
        timeCell0.Format.Alignment = ParagraphAlignment.Left;

        var timeCell1 = row.Cells[1];
        timeCell1.AddParagraph("Programação da reunião do meio de semana");
        timeCell1.Format.Font = new Font("Arial", 16);
        timeCell1.Format.Font.Bold = true;
        timeCell1.Format.Alignment = ParagraphAlignment.Left;
    }

    private void WriteDateAndPresidentRow(Section section, string date = "[DATA]", string presidentName = "[Nome]", string weeklyBibleReading = "[Leitura Semanal da Bíblia]", bool addBlankLine = false)
    {
        var table = section.AddTable();
        table.AddColumn(Unit.FromCentimeter(10));
        table.AddColumn(Unit.FromCentimeter(3));
        table.AddColumn(Unit.FromCentimeter(4));

        if (addBlankLine)
        {
            var blankRow = table.AddRow();
            blankRow.HeightRule = RowHeightRule.AtLeast;
            blankRow.Height = Unit.FromCentimeter(0.5);
        }

        var row = table.AddRow();
        row.HeightRule = RowHeightRule.AtLeast;
        row.Height = Unit.FromCentimeter(1);
        row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;

        var timeCell0 = row.Cells[0];
        timeCell0.AddParagraph($"{date} | {weeklyBibleReading}");
        timeCell0.Format.Font = new Font("Arial", 11);
        timeCell0.Format.Font.Bold = true;
        timeCell0.Format.Alignment = ParagraphAlignment.Left;

        var timeCell1 = row.Cells[1];
        timeCell1.AddParagraph("Presidente:");
        timeCell1.Format.Font = new Font("Arial", 8);
        timeCell1.Format.Font.Bold = true;
        timeCell1.Format.Alignment = ParagraphAlignment.Right;

        var timeCell2 = row.Cells[2];
        timeCell2.AddParagraph(presidentName);
        timeCell2.Format.Font = new Font("Arial", 10);
        timeCell2.Format.Alignment = ParagraphAlignment.Left;
    }

    private void WriteHeaderRow(Table table, string name, Color color)
    {
        var row = table.AddRow();
        row.HeightRule = RowHeightRule.AtLeast;
        row.Height = Unit.FromCentimeter(0.6);
        row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;

        var timeCell0 = row.Cells[0];
        timeCell0.MergeRight = 1;
        timeCell0.AddParagraph(name);
        timeCell0.Format.Font = new Font("Arial", 11);
        timeCell0.Format.Font.Bold = true;
        timeCell0.Format.Font.Color = Colors.White;
        timeCell0.Shading.Color = color;
        timeCell0.Format.Alignment = ParagraphAlignment.Left;
    }

    private void AddRow(Table table, string time, string theme = "", string title = "", string name = "", short? partNumber = null)
    {
        var row = table.AddRow();
        row.HeightRule = RowHeightRule.AtLeast;
        row.Height = Unit.FromCentimeter(0.6);
        row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;

        var timeCell0 = row.Cells[0];
        timeCell0.AddParagraph(time);
        timeCell0.Style = "TimeStyle";

        if (partNumber.HasValue && !string.IsNullOrWhiteSpace(theme))
        {
            theme = theme.Trim();
            if (!theme.StartsWith($"{partNumber}."))
                theme = $"{partNumber}. {theme}";
        }
        
        var timeCell1 = row.Cells[1];
        timeCell1.AddParagraph(theme);
        timeCell1.Format.Font = new Font("Arial", 11);
        timeCell1.Format.Alignment = ParagraphAlignment.Left;

        var timeCell2 = row.Cells[2];
        timeCell2.AddParagraph(string.IsNullOrWhiteSpace(title) ? string.Empty : $"{title}:");
        timeCell2.Format.Font = new Font("Arial", 8);
        timeCell2.Format.Font.Bold = true;
        timeCell2.Format.Alignment = ParagraphAlignment.Right;

        var timeCell3 = row.Cells[3];
        timeCell3.AddParagraph(name);
        timeCell3.Format.Font = new Font("Arial", 10);
        timeCell3.Format.Alignment = ParagraphAlignment.Left;
    }

    private void DefineStyles(Document document)
    {
        Style style = document.Styles["Normal"];
        style.Font.Name = "Arial";
        style.Font.Size = 10;
        style.Font.Color = Color.FromRgb(0, 0, 0);

        style = document.Styles.AddStyle("NormalBold", "Normal");
        style.Font.Bold = true;

        style = document.Styles.AddStyle("Heading1", "Normal");
        style.Font.Size = 16;
        style.Font.Bold = true;
        style.ParagraphFormat.SpaceAfter = "0.5cm";

        style = document.Styles.AddStyle("Heading2", "Normal");
        style.Font.Size = 14;
        style.Font.Bold = true;
        style.ParagraphFormat.SpaceAfter = "0.8cm";

        style = document.Styles.AddStyle("SectionHeader", "Normal");
        style.Font.Size = 10;
        style.Font.Bold = true;
        style.Font.Color = Color.FromRgb(255, 255, 255);
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        style = document.Styles.AddStyle("TimeStyle", "Normal");
        style.Font = new Font("Arial", 10);
        style.Font.Color = Color.FromRgb(87, 90, 93);
        style.Font.Bold = true;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
    }
}
