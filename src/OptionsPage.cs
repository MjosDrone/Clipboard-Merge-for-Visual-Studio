namespace ExternalMerge.VSExtension
{
  using System.ComponentModel;
  using Microsoft.VisualStudio.Shell;

  public class OptionsPage : DialogPage
  {
    #region Properties & Fields - Public

    [Category("External Merge Tool")]
    [DisplayName("Tool Path")]
    [Description("Full path to the external merge tool executable (e.g., ...\\BCompare.exe).")]
    public string ToolPath { get; set; } = ""; // Default value

    [Category("External Merge Tool")]
    [DisplayName("Tool Arguments")]
    [Description("Command line arguments. Use {filePath1} (editor), {filePath2} (clipboard), {outputFilePath} (merged result).")]
    public string ToolArguments { get; set; } = "{filePath1} {filePath2} /savetarget={outputFilePath}"; // Example for Beyond Compare

    #endregion

    // The 3-way merge logic from your TS file can be added here if needed.
    // For simplicity, this example sticks to the 2-way diff + output file pattern.
  }
}
