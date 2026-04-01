using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using NwScheduleFormatter.Configuration;
using NwScheduleFormatter.Models;

namespace NwScheduleFormatter.Services;

public class FileService
{
    public void CreateStudantDesignationDocuments(List<Meeting> meetings, string outputPath)
    {
        // 1. Cria um novo documento MigraDoc
        var document = new Document();
        document.Info.Title = "Programação da Reunião do Meio de Semana";
        document.Info.Author = ApplicationSettings.CONGREGATION_NAME;

        // 2. Define os estilos padrão do documento
        DefineStyles(document);

        // 3. Adiciona uma nova seção ao documento
        Section section = document.AddSection();
        section.PageSetup.PageFormat = PageFormat.A4;
        section.PageSetup.LeftMargin = "1.0cm";
        section.PageSetup.RightMargin = "1.0cm";
        section.PageSetup.TopMargin = "1.0cm";
        section.PageSetup.BottomMargin = "1.0cm";

        var table = section.AddTable();
        table.AddColumn(Unit.FromCentimeter(9.2));
        table.AddColumn(Unit.FromCentimeter(9.2));
        table.Rows.Alignment = RowAlignment.Center;
        table.TopPadding = 10;
        table.BottomPadding = 10;
        table.LeftPadding = 10;
        table.RightPadding = 10;

        var designationCards = new List<DesignationCard>();
        foreach (var meeting in meetings)
        {
            var biblieReading = new Presentation { Speaker = meeting.TreasuresFromGodsWord.BibleReading };
            designationCards.Add(new DesignationCard(meeting.Date.AddDays((int)ApplicationSettings.MeetingDay), biblieReading, 3));
            designationCards.Add(new DesignationCard(meeting.Date.AddDays((int)ApplicationSettings.MeetingDay), meeting.Apply1, 4));
            designationCards.Add(new DesignationCard(meeting.Date.AddDays((int)ApplicationSettings.MeetingDay), meeting.Apply2, 5));
            designationCards.Add(new DesignationCard(meeting.Date.AddDays((int)ApplicationSettings.MeetingDay), meeting.Apply3, 6));
            if (!string.IsNullOrWhiteSpace(meeting.Apply4.Speaker))
                designationCards.Add(new DesignationCard(meeting.Date.AddDays((int)ApplicationSettings.MeetingDay), meeting.Apply4, 7));
        }

        for (var i = 0; i < designationCards.Count; i = i + 2)
        {
            AddRow(table, designationCards[i], i + 1 < designationCards.Count ? designationCards[i + 1] : null);
        }

        // 10. Renderiza o documento MigraDoc para um documento PDFsharp
        PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
        pdfRenderer.Document = document;
        pdfRenderer.RenderDocument();

        // 11. Salva o documento PDFsharp
        pdfRenderer.PdfDocument.Save(outputPath);
    }

    /// <summary>
    /// Define estilos personalizados para o documento.
    /// </summary>
    private void DefineStyles(Document document)
    {
        // Estilo Normal
        MigraDoc.DocumentObjectModel.Style style = document.Styles["Normal"];
        style.Font.Name = "Arial";
        style.Font.Size = 12;
        style.Font.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(0, 0, 0); // Preto

        // Label style (para os rótulos "Nome:")
        style = document.Styles.AddStyle("Label", "Normal");
        style.Font = new Font("Arial", 12);
        style.Font.Bold = true;

        // Estilo para Títulos (Congregação)
        style = document.Styles.AddStyle("Heading1", "Normal");
        style.Font = new Font("Arial", 18);
        style.Font.Bold = true;
        style.ParagraphFormat.SpaceAfter = "0.5cm";
        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Estilo para Observações em negrito
        style = document.Styles.AddStyle("ObservationBold", "Normal");
        style.Font = new Font("Arial", 10.5);
        style.Font.Bold = true;

        // Estilo para Observações
        style = document.Styles.AddStyle("Observation", "Normal");
        style.Font = new Font("Arial", 10.5);
    }

    private void AddRow(Table table, DesignationCard part1, DesignationCard part2 = null)
    {
        var row = table.AddRow();
        row.HeightRule = RowHeightRule.AtLeast;
        row.Height = Unit.FromCentimeter(10);
        row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Top;
        row.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(205, 205, 205);

        var timeCell0 = row.Cells[0];
        timeCell0.Elements.Add(CreateDesignationCard(part1));

        if (part2?.Presentation?.Speaker == null)
            return;

        var timeCell1 = row.Cells[1];
        timeCell1.Elements.Add(CreateDesignationCard(part2));
    }

    public Table CreateDesignationCard(DesignationCard designationCard)
    {
        var designationTable = new Table();
        designationTable.AddColumn(Unit.FromCentimeter(8.5));
        designationTable.Rows.Height = Unit.FromCentimeter(1);
        designationTable.Rows.HeightRule = RowHeightRule.AtLeast;

        var titleRow = designationTable.AddRow();
        titleRow.Cells[0].AddParagraph("DESIGNAÇÃO PARA A REUNIÃO NOSSA VIDA E MINISTÉRIO CRISTÃO");
        titleRow.Format.Font = new Font("Arial", 11);
        titleRow.Format.Font.Bold = true;
        titleRow.Style = "Heading1";
        titleRow.Format.Alignment = ParagraphAlignment.Center;

        var nameRow = designationTable.AddRow();
        var pName = nameRow.Cells[0].AddParagraph();
        pName.AddText("Nome: ");
        pName.Style = "Label";
        var name = pName.AddFormattedText(designationCard.Presentation.Speaker);
        name.Style = "Normal";

        var assistantRow = designationTable.AddRow();
        var pAssistant = assistantRow.Cells[0].AddParagraph();
        pAssistant.AddText("Ajudante: ");
        pAssistant.Style = "Label";
        var helper = pAssistant.AddFormattedText(designationCard.Presentation.Assistant ?? string.Empty);
        helper.Style = "Normal";

        var dateRow = designationTable.AddRow();
        var pDate = dateRow.Cells[0].AddParagraph();
        pDate.AddText("Data: ");
        pDate.Style = "Label";
        var date = pDate.AddFormattedText(designationCard.Date.ToString("dd/MM/yyyy"));
        date.Style = "Normal";

        var partNumberRow = designationTable.AddRow();
        var pPartNumber = partNumberRow.Cells[0].AddParagraph();
        pPartNumber.AddText("Número da parte: ");
        pPartNumber.Style = "Label";
        var partNumber = pPartNumber.AddFormattedText(designationCard.PartNumber.ToString());
        partNumber.Style = "Normal";

        designationTable.AddRow().Height = Unit.FromCentimeter(0.5);

        var roomRow = designationTable.AddRow();
        var pRoom = roomRow.Cells[0].AddParagraph();
        pRoom.AddText("Local: ");
        pRoom.Style = "Label";
        var room = pRoom.AddFormattedText("Salão principal");
        room.Style = "Normal";

        designationTable.AddRow().Height = Unit.FromCentimeter(0.5);

        var observationRow = designationTable.AddRow();
        var pObservation = observationRow.Cells[0].AddParagraph();
        pObservation.AddText("Observação para o estudante: ");
        pObservation.Style = "ObservationBold";
        var observation = pObservation.AddFormattedText("A lição e a fonte de matéria para a sua designação estão na Apostila da Reunião Vida e Ministério. Veja as instruções para a parte que estão nas Instruções para a Reunião Nossa Vida e Ministério Cristão (S-38).");
        observation.Style = "Observation";
        pObservation.Format.Alignment = ParagraphAlignment.Justify;


        var formVersionRow = designationTable.AddRow();
        var pFormVersion = formVersionRow.Cells[0].AddParagraph("S-89-T    11/23");
        pFormVersion.Format.Font = new Font("Arial", 9);
        formVersionRow.Format.Alignment = ParagraphAlignment.Left;
        formVersionRow.VerticalAlignment = VerticalAlignment.Bottom;

        return designationTable;
    }
}
