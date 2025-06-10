namespace ExternalMerge.VSExtension
{
  using System;
  using System.ComponentModel.Design;
  using System.Diagnostics;
  using System.IO;
  using System.Threading.Tasks;
  using System.Windows;
  using EnvDTE;
  using EnvDTE80;
  using Microsoft.VisualStudio.Shell;
  using Microsoft.VisualStudio.Shell.Interop;
  using Process = System.Diagnostics.Process;
  using Task = System.Threading.Tasks.Task;

  /// <summary>Command handler</summary>
  internal sealed class MergeWithClipboardCommand
  {
    #region Constants & Statics

    /// <summary>Command ID.</summary>
    public const int CommandId = 0x0100;

    /// <summary>Command menu group (command set GUID).</summary>
    public static readonly Guid CommandSet = new Guid("aeb97d7a-3766-41b4-bfb8-67cdc0360043");

    /// <summary>Gets the instance of the command.</summary>
    public static MergeWithClipboardCommand Instance { get; private set; }

    #endregion

    #region Properties & Fields - Non-Public

    /// <summary>VS Package that provides this command, not null.</summary>
    private readonly AsyncPackage package;

    private readonly DTE2 dte;

    /// <summary>Gets the service provider from the owner package.</summary>
    private IAsyncServiceProvider ServiceProvider
    {
      get { return package; }
    }

    #endregion

    #region Constructors

    /// <summary>
    ///   Initializes a new instance of the <see cref="MergeWithClipboardCommand" /> class.
    ///   Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    /// <param name="commandService">Command service to add command to, not null.</param>
    private MergeWithClipboardCommand(AsyncPackage package, OleMenuCommandService commandService, DTE2 dte)
    {
      this.package   = package ?? throw new ArgumentNullException(nameof(package));
      this.dte       = dte ?? throw new ArgumentNullException(nameof(dte));
      commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

      var menuCommandID = new CommandID(CommandSet, CommandId);
      var menuItem      = new MenuCommand(Execute, menuCommandID);
      commandService.AddCommand(menuItem);
    }

    #endregion

    #region Methods

    /// <summary>Initializes the singleton instance of the command.</summary>
    /// <param name="package">Owner package, not null.</param>
    public static async Task InitializeAsync(AsyncPackage package)
    {
      // Switch to the main thread - the call to AddCommand in MergeWithClipboardCommand2's constructor requires
      // the UI thread.
      await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

      var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
      var dte            = await package.GetServiceAsync(typeof(DTE)) as DTE2;

      _ = new MergeWithClipboardCommand(package, commandService, dte);
    }

    /// <summary>
    ///   The event handler for the command. It's `void` as required, but offloads all work to
    ///   a JTF-managed task for safety.
    /// </summary>
    private void Execute(object sender, EventArgs e)
    {
      // BEST PRACTICE: Use the JTF to run the async logic.
      // This prevents the `async void` anti-pattern and handles exceptions correctly.
      package.JoinableTaskFactory.RunAsync(async () =>
      {
        string tempFile1Path  = null;
        string tempFile2Path  = null;
        string outputFilePath = null;

        try
        {
          // We are likely on a background thread here, so switch to main to get DTE objects
          await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

          Document activeDocument = dte.ActiveDocument;
          if (activeDocument == null)
          {
            ShowMessageBox("Error", "No active document.");
            return;
          }

          TextDocument textDocument           = activeDocument.Object("TextDocument") as TextDocument;
          EditPoint    startPoint             = textDocument.StartPoint.CreateEditPoint();
          string       activeFileContent      = startPoint.GetText(textDocument.EndPoint);
          string       activeFilePathOriginal = activeDocument.FullName;

          // Getting clipboard text must be on the UI thread.
          string clipboardContent = Clipboard.GetText();
          if (string.IsNullOrEmpty(clipboardContent))
          {
            ShowMessageBox("Info", "Clipboard is empty.");
            return;
          }

          OptionsPage options               = (OptionsPage)package.GetDialogPage(typeof(OptionsPage));
          string      toolPath              = options.ToolPath;
          string      toolArgumentsTemplate = options.ToolArguments;

          if (string.IsNullOrEmpty(toolPath))
          {
            ShowMessageBox("Error", "External merge tool path is not configured. Go to Tools > Options > External Merge.");
            return;
          }

          // --- Start of I/O and Process logic, can be on a background thread ---

          string tempDir         = Path.GetTempPath();
          string timestamp       = DateTime.Now.Ticks.ToString();
          string originalFileExt = Path.GetExtension(activeFilePathOriginal);

          tempFile1Path  = Path.Combine(tempDir, $"vs-merge-{timestamp}-editor{originalFileExt}");
          tempFile2Path  = Path.Combine(tempDir, $"vs-merge-{timestamp}-clipboard{originalFileExt}");
          outputFilePath = Path.Combine(tempDir, $"vs-merge-{timestamp}-output{originalFileExt}");

          string finalOutputToRead = outputFilePath;

          // BEST PRACTICE: Use Task.Run to execute synchronous, blocking I/O on a background thread.
          await Task.Run(() => File.WriteAllText(tempFile1Path, activeFileContent));
          await Task.Run(() => File.WriteAllText(tempFile2Path, clipboardContent));
          await Task.Run(() => File.WriteAllText(outputFilePath, activeFileContent));

          string finalArguments = toolArgumentsTemplate
                                  .Replace("{filePath1}", $"\"{tempFile1Path}\"")
                                  .Replace("{filePath2}", $"\"{tempFile2Path}\"")
                                  .Replace("{outputFilePath}", $"\"{outputFilePath}\"");

          if (!toolArgumentsTemplate.Contains("{outputFilePath}"))
            finalOutputToRead = tempFile1Path;

          ProcessStartInfo startInfo = new ProcessStartInfo
          {
            FileName        = toolPath,
            Arguments       = finalArguments,
            UseShellExecute = false
          };

          using (Process process = new Process { StartInfo = startInfo })
          {
            process.Start();
            await Task.Run(() => process.WaitForExit()); // Correctly wait on a background thread
          }

          // BEST PRACTICE: Read file on a background thread.
          string mergedContent = await Task.Run(() => File.ReadAllText(finalOutputToRead));

          // --- Back to the UI thread to update the editor ---
          await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

          if (mergedContent != activeFileContent)
          {
            startPoint.ReplaceText(textDocument.EndPoint, mergedContent, (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
            //ShowMessageBox("Success", "Merge applied to the active editor.");
          }
        }
        catch (Exception ex)
        {
          // Make sure to switch to the UI thread before showing a message box
          await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
          ShowMessageBox("Error", $"Error during merge process: {ex.Message}");
        }
        finally
        {
          // Cleanup is safe to do on a background thread
          await Task.Run(() =>
          {
            if (File.Exists(tempFile1Path)) File.Delete(tempFile1Path);
            if (File.Exists(tempFile2Path)) File.Delete(tempFile2Path);
            if (File.Exists(outputFilePath)) File.Delete(outputFilePath);
          });
        }
      });
    }

    /// <summary>Helper method to show a message box, ensuring it's called on the UI thread.</summary>
    private void ShowMessageBox(string title, string message)
    {
      // This method should be called after ensuring we are on the UI thread.
      ThreadHelper.ThrowIfNotOnUIThread();
      VsShellUtilities.ShowMessageBox(
        package,
        message,
        title,
        OLEMSGICON.OLEMSGICON_INFO,
        OLEMSGBUTTON.OLEMSGBUTTON_OK,
        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
    }

    #endregion
  }
}
