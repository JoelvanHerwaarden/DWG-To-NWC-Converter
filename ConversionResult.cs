namespace DWGToNWCConverter;

public sealed class ConversionResult
{
    public required int TotalFiles { get; init; }

    public required int ConvertedFiles { get; init; }

    public required IReadOnlyList<string> Messages { get; init; }
}
