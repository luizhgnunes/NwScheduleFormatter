using System;
using NwScheduleFormatter.Models;

namespace NwScheduleFormatter.Services.Interfaces;

/// <summary>
/// Interface for PDF generation services related to meeting schedules.
/// </summary>
public interface IPdfGenerationService
{
    /// <summary>
    /// Sets the context for PDF generation (year and month).
    /// </summary>
    void SetContext(int year, int month);

    /// <summary>
    /// Generates a PDF document for the meeting schedule.
    /// </summary>
    Task GenerateMeetingSchedulePdfAsync(string outputDirectory, List<Meeting> meetings, IProgress<int> progress, int progressSlice);

    /// <summary>
    /// Generates a PDF document for student designations.
    /// </summary>
    Task GenerateStudentDesignationsPdfAsync(string outputDirectory, List<Meeting> meetings, IProgress<int> progress, int progressSlice);
}
