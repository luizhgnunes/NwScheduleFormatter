using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using NwScheduleFormatter.Configuration;
using NwScheduleFormatter.Data;
using NwScheduleFormatter.Enumerations;
using NwScheduleFormatter.Helpers;
using NwScheduleFormatter.Models;
using NwScheduleFormatter.Services;
using PdfSharp.Fonts;
using System.Windows;
using Colors = MigraDoc.DocumentObjectModel.Colors;
using Section = MigraDoc.DocumentObjectModel.Section;
using Style = MigraDoc.DocumentObjectModel.Style;
using Table = MigraDoc.DocumentObjectModel.Tables.Table;

namespace NwScheduleFormatter;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        PopulateYearComboBox();
        PopulateMonthComboBox();
    }

    private void PopulateYearComboBox()
    {
        var currentYear = DateTime.Now.Year;

        for (int i = 0; i <= 1; i++)
        {
            YearComboBox.Items.Add(currentYear + i);
        }
        YearComboBox.SelectedIndex = 0;
    }

    private void PopulateMonthComboBox()
    {
        for (int i = 1; i <= 12; i++)
        {
            MonthComboBox.Items.Add(i);
        }
        var currentMonth = DateTime.Now.Month;
        MonthComboBox.SelectedIndex = currentMonth - 1;
    }

    private async void GenerateButton_Click(object sender, RoutedEventArgs e)
    {
        await GenerateMeetingSchedulePdf(
            @"C:\Destination\Path");
    }

    public async Task GenerateMeetingSchedulePdf(string outputPath)
    {
        var nwCsvRecords = NwCsvReader.ReadFromCsv(
            InputFileTextBox.Text,
            (int)YearComboBox.SelectedItem, (int)MonthComboBox.SelectedItem);

        var jwMeetingReader = new JwMeetingReader();
        var meetings = new List<Meeting>();
        foreach (var record in nwCsvRecords)
        {
            var meeting = meetings.Find(x => x.Date == record.Date);
            if (meeting == null)
            {
                meeting = new Meeting(record.Date);

                var jwWebsiteMeeting = await jwMeetingReader.ReadFromJwWebsiteAsync(meeting.Date);
                meeting.InitialSong = jwWebsiteMeeting.InitialSong;
                meeting.MiddleSong = jwWebsiteMeeting.MiddleSong;
                meeting.FinalSong = jwWebsiteMeeting.FinalSong;
                meeting.Apply1.DurationMinutes = jwWebsiteMeeting.Apply1DurationMinutes;
                meeting.Apply2.DurationMinutes = jwWebsiteMeeting.Apply2DurationMinutes;
                meeting.Apply3.DurationMinutes = jwWebsiteMeeting.Apply3DurationMinutes;
                meeting.Apply4.DurationMinutes = jwWebsiteMeeting.Apply4DurationMinutes;
                meeting.LivingAsChristhians1.DurationMinutes = jwWebsiteMeeting.LivingAsChristhians1DurationMinutes;
                meeting.LivingAsChristhians2.DurationMinutes = jwWebsiteMeeting.LivingAsChristhians2DurationMinutes;

                meetings.Add(meeting);
            }

            var partType = Enum.Parse<PartType>(record.PartType);
            switch (partType)
            {
                case PartType.Chairman:
                    meeting.Chairman = record.Person;
                    break;
                case PartType.OpeningPrayer:
                    meeting.OpeningPrayer = record.Person;
                    break;
                case PartType.ClosingPrayer:
                    meeting.ClosingPrayer = record.Person;
                    break;

                case PartType.TreasuresTalk:
                    meeting.TreasuresFromGodsWord.Treasures
                            = new Speech
                            {
                                Theme = record.Assignment,
                                Speaker = record.Person,
                                DurationMinutes = 10
                            };
                    break;
                case PartType.SpiritualGems:
                    meeting.TreasuresFromGodsWord.SpiritualGems = record.Person;
                    break;
                case PartType.BibleReading:
                    meeting.TreasuresFromGodsWord.BibleReading = record.Person;
                    break;

                case PartType.Apply1:
                    meeting.Apply1.Theme = record.Assignment;
                    meeting.Apply1.Speaker = record.Person;
                    break;
                case PartType.Apply2:
                    meeting.Apply2.Theme = record.Assignment;
                    meeting.Apply2.Speaker = record.Person;
                    break;
                case PartType.Apply3:
                    meeting.Apply3.Theme = record.Assignment;
                    meeting.Apply3.Speaker = record.Person;
                    break;
                case PartType.Apply4:
                    meeting.Apply4.Theme = record.Assignment;
                    meeting.Apply4.Speaker = record.Person;
                    break;
                case PartType.Apply1Assistant:
                    meeting.Apply1.Assistant = record.Person;
                    break;
                case PartType.Apply2Assistant:
                    meeting.Apply2.Assistant = record.Person;
                    break;
                case PartType.Apply3Assistant:
                    meeting.Apply3.Assistant = record.Person;
                    break;
                case PartType.Apply4Assistant:
                    meeting.Apply4.Assistant = record.Person;
                    break;

                case PartType.Living1:
                    meeting.LivingAsChristhians1.Theme = record.Assignment;
                    meeting.LivingAsChristhians1.Speaker = record.Person;
                    break;
                case PartType.Living2:
                    meeting.LivingAsChristhians2.Theme = record.Assignment;
                    meeting.LivingAsChristhians2.Speaker = record.Person;
                    break;
                case PartType.CBS:
                    meeting.CongregationBibleStudy.Speaker = record.Person;
                    break;
                case PartType.CBSReader:
                    meeting.CongregationBibleStudy.Reader = record.Person;
                    break;
            }
        }

        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        // 1. Cria um novo documento MigraDoc
        var document = new Document();
        document.Info.Title = "Programação da Reunião do Meio de Semana";
        document.Info.Author = ApplicationSettings.CONGREGATION_NAME;

        // 2. Define os estilos padrão do documento
        DefineStyles(document);

        // 3. Adiciona uma nova seção ao documento
        Section section = document.AddSection();
        section.PageSetup.PageFormat = PageFormat.A4;
        section.PageSetup.LeftMargin = "2.0cm";
        section.PageSetup.RightMargin = "2.0cm";
        section.PageSetup.TopMargin = "1.0cm";
        section.PageSetup.BottomMargin = "1.0cm";

        for (var i = 0; i < meetings.Count; i++)
        {
            if (i > 1 && (i + 1) % 2 > 0)
                section.AddPageBreak();

            var addBlankRow = false;
            if ((i + 1) % 2 > 0)
                WritePageHeader(section);
            else
                addBlankRow = true;

            WriteDateAndPresidentRow(section, DateHelper.GetWeekInFull(meetings[i].Date).ToUpper(), meetings[i].Chairman, addBlankRow);

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

            WriteHeaderRow(table, "TESOUROS DA PALAVRA DE DEUS", MigraDoc.DocumentObjectModel.Color.FromRgb(87, 90, 93));
            time = time.AddMinutes(1);
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].TreasuresFromGodsWord.Treasures.Theme} (10 min)", "", meetings[i].TreasuresFromGodsWord.Treasures.Speaker);
            time = time.AddMinutes(10);
            AddRow(table, time.ToString("HH:mm"), "2. Joias espirituais (10 min)", "", meetings[i].TreasuresFromGodsWord.SpiritualGems);
            time = time.AddMinutes(10);
            AddRow(table, time.ToString("HH:mm"), "3. Leitura da Bíblia (4 min)", "", meetings[i].TreasuresFromGodsWord.BibleReading);

            WriteHeaderRow(table, "FAÇA SEU MELHOR NO MINISTÉRIO", MigraDoc.DocumentObjectModel.Color.FromRgb(190, 137, 0));
            
            time = time.AddMinutes(5);
            var apply1Assistant = string.IsNullOrWhiteSpace(meetings[i].Apply1.Assistant + 1) ? string.Empty : $" / {meetings[i].Apply1.Assistant}";
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].Apply1.Theme} ({meetings[i].Apply1.DurationMinutes} min)", "", $"{meetings[i].Apply1.Speaker}{apply1Assistant}");
            
            time = time.AddMinutes((int)meetings[i].Apply1.DurationMinutes + 1);
            var apply2Assistant = string.IsNullOrWhiteSpace(meetings[i].Apply2.Assistant) ? string.Empty : $" / {meetings[i].Apply2.Assistant}";
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].Apply2.Theme} ({meetings[i].Apply2.DurationMinutes} min)", "", $"{meetings[i].Apply2.Speaker}{apply2Assistant}");
            
            time = time.AddMinutes((int)meetings[i].Apply2.DurationMinutes + 1);
            var apply3Assistant = string.IsNullOrWhiteSpace(meetings[i].Apply3.Assistant) ? string.Empty : $" / {meetings[i].Apply3.Assistant}";
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].Apply3.Theme} ({meetings[i].Apply3.DurationMinutes} min)", "", $"{meetings[i].Apply3.Speaker}{apply3Assistant}");
            
            short nextPartNumber = 7;
            if (!string.IsNullOrWhiteSpace(meetings[i].Apply4.Speaker))
            {
                time = time.AddMinutes((int)meetings[i].Apply3.DurationMinutes);
                var apply4Assistant = string.IsNullOrWhiteSpace(meetings[i].Apply4.Assistant) ? string.Empty : $" / {meetings[i].Apply4.Assistant}";
                AddRow(table, time.ToString("HH:mm"), $"{meetings[i].Apply4.Theme} ({meetings[i].Apply4.DurationMinutes} min)", "", $"{meetings[i].Apply4.Speaker}{apply4Assistant}");
                nextPartNumber++;
            }

            WriteHeaderRow(table, "NOSSA VIDA CRISTÃ", MigraDoc.DocumentObjectModel.Color.FromRgb(126, 0, 36));
            int previousPartMinutes = string.IsNullOrWhiteSpace(meetings[i].Apply4.Speaker) ? (int)meetings[i].Apply3.DurationMinutes : (int)meetings[i].Apply4.DurationMinutes;

            time = time.AddMinutes(previousPartMinutes + 1);
            AddRow(table, time.ToString("HH:mm"), $"• Cântico {meetings[i].MiddleSong.Number}");
            time = time.AddMinutes(5);
            AddRow(table, time.ToString("HH:mm"), $"{meetings[i].LivingAsChristhians1.Theme} ({meetings[i].LivingAsChristhians1.DurationMinutes} min)", "", $"{meetings[i].LivingAsChristhians1.Speaker}");
            nextPartNumber++;
            if (!string.IsNullOrWhiteSpace(meetings[i].LivingAsChristhians2.Speaker))
            {
                time = time.AddMinutes((int)meetings[i].LivingAsChristhians1.DurationMinutes);
                AddRow(table, time.ToString("HH:mm"), $"{meetings[i].LivingAsChristhians2.Theme} ({meetings[i].LivingAsChristhians2.DurationMinutes} min)", "", $"{meetings[i].LivingAsChristhians2.Speaker}");
                nextPartNumber++;
            }

            previousPartMinutes = string.IsNullOrWhiteSpace(meetings[i].LivingAsChristhians2.Speaker) ? (int)meetings[i].LivingAsChristhians1.DurationMinutes : (int)meetings[i].LivingAsChristhians2.DurationMinutes;
            time = time.AddMinutes(previousPartMinutes);
            AddRow(table, time.ToString("HH:mm"), $"{nextPartNumber}. Estudo bíblico de congregação (30 min)", "Dirigente/leitor", $"{meetings[i].CongregationBibleStudy.Speaker} / {meetings[i].CongregationBibleStudy.Reader}");
            time = time.AddMinutes(30);
            AddRow(table, time.ToString("HH:mm"), "• Comentários finais (3 min)");
            AddRow(table, time.AddMinutes(3).ToString("HH:mm"), $"• Cântico {meetings[i].FinalSong.Number}", "Oração", $"{meetings[i].ClosingPrayer}");
        }


        // 10. Renderiza o documento MigraDoc para um documento PDFsharp
        PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
        pdfRenderer.Document = document;
        pdfRenderer.RenderDocument();

        // 11. Salva o documento PDFsharp
        pdfRenderer.PdfDocument.Save(outputPath);
        pdfRenderer.PdfDocument.Close();

        var fileService = new FileService();
        fileService.CreateStudantDesignationDocuments(meetings, outputPath.Replace(".pdf", "_Designações.pdf"));
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

    private void WriteDateAndPresidentRow(Section section, string date = "[DATA]", string presidentName = "[Nome]", bool addBlankLine = false)
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
        timeCell0.AddParagraph($"{date} | LEITURA SEMANAL DA BÍBLIA");
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
        timeCell2.Format.Font = new Font("Arial", 11);
        timeCell2.Format.Alignment = ParagraphAlignment.Left;
    }

    private void WriteHeaderRow(Table table, string name, MigraDoc.DocumentObjectModel.Color color)
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

    private void AddRow(Table table, string time, string theme = "", string title = "", string name = "")
    {
        var row = table.AddRow();
        row.HeightRule = RowHeightRule.AtLeast;
        row.Height = Unit.FromCentimeter(0.6);
        row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;

        var timeCell0 = row.Cells[0];
        timeCell0.AddParagraph(time);
        timeCell0.Style = "TimeStyle";

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
        timeCell3.Format.Font = new Font("Arial", 11);
        timeCell3.Format.Alignment = ParagraphAlignment.Left;
    }

    /// <summary>
    /// Define estilos personalizados para o documento.
    /// </summary>
    private void DefineStyles(Document document)
    {
        // Estilo Normal
        Style style = document.Styles["Normal"];
        style.Font.Name = "Arial";
        style.Font.Size = 10;
        style.Font.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(0, 0, 0); // Preto

        // Estilo Normal Bold (para a linha da data)
        style = document.Styles.AddStyle("NormalBold", "Normal");
        style.Font.Bold = true;

        // Estilo para Títulos (Congregação)
        style = document.Styles.AddStyle("Heading1", "Normal");
        style.Font.Size = 16;
        style.Font.Bold = true;
        style.ParagraphFormat.SpaceAfter = "0.5cm";

        // Estilo para Subtítulos (Programação)
        style = document.Styles.AddStyle("Heading2", "Normal");
        style.Font.Size = 14;
        style.Font.Bold = true;
        style.ParagraphFormat.SpaceAfter = "0.8cm";

        // Estilo para cabeçalhos de seção na tabela
        style = document.Styles.AddStyle("SectionHeader", "Normal");
        style.Font.Size = 10;
        style.Font.Bold = true;
        style.Font.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(255, 255, 255); // Branco
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Time style
        style = document.Styles.AddStyle("TimeStyle", "Normal");
        style.Font = new Font("Arial", 10);
        style.Font.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(87, 90, 93);
        style.Font.Bold = true;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
    }
}
