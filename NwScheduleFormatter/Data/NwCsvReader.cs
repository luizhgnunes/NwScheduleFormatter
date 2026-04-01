using NwScheduleFormatter.Models;
using System.Globalization;
using System.IO;

namespace NwScheduleFormatter.Data;

public static class NwCsvReader
{
    /// <summary>
    /// Lê um arquivo CSV e retorna uma lista de objetos NwCsv.
    /// Suporta campos entre aspas com vírgulas internas (ex: "1. Seja uma muralha, não uma porta").
    /// </summary>
    /// <param name="csvPath">Caminho completo do arquivo CSV.</param>
    /// <returns>Lista de NwCsv com os dados do arquivo.</returns>
    public static IEnumerable<NwCsv> ReadFromCsv(string csvPath, int year, int month)
    {
        var result = new List<NwCsv>();

        if (!File.Exists(csvPath))
            throw new FileNotFoundException($"Arquivo CSV não encontrado: {csvPath}");

        var lines = File.ReadAllLines(csvPath);

        // Ignora a linha de cabeçalho (índice 0)
        for (int i = 1; i < lines.Length; i++)
        {
            string line = lines[i].Trim();

            if (string.IsNullOrWhiteSpace(line))
                continue;

            int yearInLine = int.Parse(line.Substring(0, 4));
            int monthInLine = int.Parse(line.Substring(5, 2));

            if (yearInLine != year || monthInLine != month)
                continue;

            List<string> fields = ParseCsvLine(line);

            // Garante que a linha possui exatamente 5 colunas
            if (fields.Count < 5)
                throw new FormatException(
                    $"Linha {i + 1} possui apenas {fields.Count} coluna(s). " +
                    $"Esperado: 5. Conteúdo: \"{line}\"");

            var record = new NwCsv
            {
                Date = DateOnly.ParseExact(fields[0].Trim(), "yyyy/MM/dd",
                                 CultureInfo.InvariantCulture),
                Person = fields[1].Trim(),
                PartType = fields[2].Trim(),
                Assignment = fields[3].Trim(),
                School = fields[4].Trim()
            };

            result.Add(record);
        }

        return result.OrderBy(x => x.Date).ToList();
    }

    /// <summary>
    /// Faz o parse de uma linha CSV respeitando campos entre aspas duplas.
    /// Exemplo: 2025/11/17,Edson,"1. Seja uma muralha, não uma porta",TreasuresTalk,1
    /// </summary>
    private static List<string> ParseCsvLine(string line)
    {
        var fields = new List<string>();
        var currentField = new System.Text.StringBuilder();
        bool insideQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (c == '"')
            {
                if (insideQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    // Aspas duplas escapadas ("") → representa uma aspa literal
                    currentField.Append('"');
                    i++; // Pula o próximo "
                }
                else
                {
                    // Alterna entre dentro/fora das aspas
                    insideQuotes = !insideQuotes;
                }
            }
            else if (c == ',' && !insideQuotes)
            {
                // Fim do campo atual
                fields.Add(currentField.ToString());
                currentField.Clear();
            }
            else
            {
                currentField.Append(c);
            }
        }

        // Adiciona o último campo
        fields.Add(currentField.ToString());

        return fields;
    }
}