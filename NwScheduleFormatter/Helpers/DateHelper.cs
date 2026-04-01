namespace NwScheduleFormatter.Helpers;

public static class DateHelper
{
    public static string GetWeekInFull(DateOnly mondayDate)
    {
        var monthName = GetMonthName(mondayDate.Month);
        var weekEndDate = mondayDate.AddDays(6);
        var isWeekEndsNextMonth = weekEndDate.Month != mondayDate.Month;

        if (isWeekEndsNextMonth)
        {
            var nextMonthName = GetMonthName(weekEndDate.Month);
            return $"{mondayDate.Day} de {monthName}–{weekEndDate.Day} de {nextMonthName}";
        }

        return $"{mondayDate.Day}-{mondayDate.Day + 6} de {monthName}";
    }

    private static string GetMonthName(int month)
    {
        return month switch
        {
            1 => "janeiro",
            2 => "fevereiro",
            3 => "março",
            4 => "abril",
            5 => "maio",
            6 => "junho",
            7 => "julho",
            8 => "agosto",
            9 => "setembro",
            10 => "outubro",
            11 => "novembro",
            12 => "dezembro",
            _ => throw new ArgumentOutOfRangeException(nameof(month), $"Mês inválido: {month}")
        };
    }
}
