using BEPDesigner.WinUI.Models;

namespace BEPDesigner.WinUI.Services;

public static class BepDefaults
{
    public static IReadOnlyList<ClashSessionState> DefaultClashSessions() =>
    [
        new() { Name = "00_Grid and Levels Check", DisciplinePair = "ARCH x STRC", Keep = true },
        new() { Name = "01_Hard Clash Core", DisciplinePair = "ARCH x MEP", Keep = true },
        new() { Name = "02_Soft Clash Clearances", DisciplinePair = "MEP x STRC", Keep = true },
        new() { Name = "03_MEP Major Runs", DisciplinePair = "MECH x ELEC", Keep = true },
        new() { Name = "04_Envelope Penetrations", DisciplinePair = "ARCH x MECH", Keep = true },
        new() { Name = "05_Weekly Delta Review", DisciplinePair = "All Disciplines", Keep = true }
    ];
}
