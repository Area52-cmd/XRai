using XRai.Demo.MacroGuard.Models;

namespace XRai.Demo.MacroGuard.Scanner;

public interface IScannerRule
{
    string Name { get; }
    bool RequiresStrictMode { get; }
    List<VbaIssue> Check(VbaModuleInfo module);
}
