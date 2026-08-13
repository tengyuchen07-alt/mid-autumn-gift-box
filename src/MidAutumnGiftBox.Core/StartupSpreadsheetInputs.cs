namespace MidAutumnGiftBox.Core;

public sealed record StartupSpreadsheetInputs(
    string RootDirectory,
    string? ErpPath,
    string? ManualPath,
    string? PosPath,
    IReadOnlyList<string> MissingFileNames)
{
    public const string ErpFileName = "ERP產出數量.xlsx";
    public const string ManualFileName = "蛋黃酥-數量.xlsx";
    public const string PosFileName = "百貨專櫃門市預購報表.xlsx";

    public static StartupSpreadsheetInputs Resolve(string rootDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rootDirectory);
        var root = Path.GetFullPath(rootDirectory);
        var missing = new List<string>();

        string? ExistingPath(string fileName)
        {
            var path = Path.Combine(root, fileName);
            if (File.Exists(path)) return path;
            missing.Add(fileName);
            return null;
        }

        return new StartupSpreadsheetInputs(
            root,
            ExistingPath(ErpFileName),
            ExistingPath(ManualFileName),
            ExistingPath(PosFileName),
            missing);
    }
}
