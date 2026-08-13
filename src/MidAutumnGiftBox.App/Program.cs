using MidAutumnGiftBox.Core;

namespace MidAutumnGiftBox.App;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        var startupInputs = args.Length == 2 &&
                            args[0].Equals("--auto-import", StringComparison.OrdinalIgnoreCase)
            ? StartupSpreadsheetInputs.Resolve(args[1])
            : null;
        Application.Run(new MainForm(startupInputs));
    }
}
