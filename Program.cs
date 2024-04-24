// See https://aka.ms/new-console-template for more information
using System.Diagnostics;
using System.Management.Automation;

Console.WriteLine("Hello, World!");
var sourcePath = args[0];
var targetPath = args[1];

var directories = Directory.GetDirectories(sourcePath, "*", SearchOption.TopDirectoryOnly);
foreach (var directory in directories)
{
    var exe = Directory.GetFiles(directory, "*.exe").FirstOrDefault();
    if (exe == null)
    {
        continue;
    }

    var info = FileVersionInfo.GetVersionInfo(exe);
    var lnkName = (string.IsNullOrEmpty(info.FileDescription) ?
        Path.GetFileNameWithoutExtension(exe) : info.FileDescription) + ".lnk";
    var lnkPath = Path.Combine(targetPath, lnkName);

    if (File.Exists(lnkPath))
    {
        File.Delete(lnkPath);
    }

    using var powerShell = PowerShell.Create();
    powerShell.AddScript($@"
$linkPath        = ""{lnkPath}""
$targetPath      = ""{exe}""
$targetDirectory = ""{directory}""
$link            = (New-Object -ComObject WScript.Shell).CreateShortcut($linkPath)
$link.TargetPath = $targetPath
$link.WorkingDirectory = $targetDirectory
$link.Save()");
    powerShell.Invoke();
}