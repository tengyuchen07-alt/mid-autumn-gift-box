namespace MidAutumnGiftBox.Core;

public static class TestWorkbookExportWorkflow
{
    public static async Task ExportNewAsync(
        string path,
        IReadOnlyList<OrderChangeEntry> automaticEntries,
        ManualOverrideStore manualOverrideStore,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(manualOverrideStore);
        var persistedState = await manualOverrideStore.LoadAsync(cancellationToken);
        ExportNew(path, automaticEntries, persistedState);
    }

    public static void ExportNew(
        string path,
        IReadOnlyList<OrderChangeEntry> automaticEntries,
        ManualOverrideState manualState)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(automaticEntries);
        ArgumentNullException.ThrowIfNull(manualState);
        var fullPath = Path.GetFullPath(path);
        if (File.Exists(fullPath))
        {
            throw new InvalidOperationException("測試匯出只能建立新檔，不能覆寫既有 Excel。");
        }

        var directory = Path.GetDirectoryName(fullPath)
                        ?? throw new InvalidOperationException("測試匯出路徑缺少資料夾。");
        Directory.CreateDirectory(directory);
        var temporaryPath = Path.Combine(
            directory,
            $".{Path.GetFileNameWithoutExtension(fullPath)}.{Guid.NewGuid():N}.tmp.xlsx");
        var effectiveEntries = ManualOverrideWorkflow.Apply(
            automaticEntries,
            manualState.Overrides);
        effectiveEntries = ManualOverrideWorkflow.ApplyAutomaticResets(
            effectiveEntries,
            manualState.PendingAutomaticResets ?? []);
        try
        {
            GiftBoxWorkbookExporter.ExportNew(
                temporaryPath,
                effectiveEntries,
                manualState.LastFactorySubmission);
            _ = GiftBoxWorkbookExporter.ReadEditableRows(temporaryPath);
            File.Move(temporaryPath, fullPath, overwrite: false);
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }
}
