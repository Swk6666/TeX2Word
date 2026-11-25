using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

[assembly: AssemblyTitle("TeX2Word CLI")]
[assembly: AssemblyProduct("TeX2Word")]
[assembly: AssemblyDescription("TeX2Word 命令行入口：将 PDF、TeX 和 Markdown 转换为含 MathType 可编辑公式的 Word 文档")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace TeX2Word {
    internal static class CliProgram {
        static int Main(string[] args) {
            // No arguments or explicit --gui launches the desktop window without blocking the shell.
            bool launchGui = args.Length == 0 ||
                args.Any(a => String.Equals(a, "--gui", StringComparison.OrdinalIgnoreCase));
            if (launchGui) {
                string exe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TeX2Word.exe");
                var forwarded = args.Where(a => !String.Equals(a, "--gui", StringComparison.OrdinalIgnoreCase))
                    .Select(CliRunner.Argument);
                Process.Start(new ProcessStartInfo {
                    FileName = exe,
                    Arguments = String.Join(" ", forwarded),
                    UseShellExecute = true
                });
                return 0;
            }
            return CliRunner.Run(args);
        }
    }
}
