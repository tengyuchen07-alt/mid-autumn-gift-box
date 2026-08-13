using MidAutumnGiftBox.Core;

namespace MidAutumnGiftBox.App;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        var hasStartupDirectory = args.Length == 2 &&
                                  (args[0].Equals("--auto-import", StringComparison.OrdinalIgnoreCase) ||
                                   args[0].Equals("--quick-run", StringComparison.OrdinalIgnoreCase));
        var startupInputs = hasStartupDirectory
            ? StartupSpreadsheetInputs.Resolve(args[1])
            : null;
        var runQuickWorkflow = hasStartupDirectory &&
                               args[0].Equals("--quick-run", StringComparison.OrdinalIgnoreCase);
        Application.Run(new MainForm(startupInputs, runQuickWorkflow));
    }
}
