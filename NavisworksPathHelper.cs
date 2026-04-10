using System.IO;

namespace DWGToNWCConverter;

public static class NavisworksPathHelper
{
    public static string GetPreferredToolPath()
    {
        var detectedPath = FindInstalledToolPath();
        return detectedPath ?? BuildToolPath(2024);
    }

    public static string? FindInstalledToolPath()
    {
        for (var year = 2030; year >= 2022; year--)
        {
            var candidatePath = BuildToolPath(year);
            if (File.Exists(candidatePath))
            {
                return candidatePath;
            }
        }

        return null;
    }

    private static string BuildToolPath(int year)
    {
        return $@"C:\Program Files\Autodesk\Navisworks Manage {year}\FiletoolsTaskRunner.exe";
    }
}
