using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace TeX2Word {
    internal static class CliRunner {
        private static readonly HashSet<string> CliFlags = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "--cli", "-h", "--help", "-V", "--version",
            "-o", "--output", "-d", "--output-dir",
            "--lang", "--date", "--force", "--verify"
        };

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(int dwProcessId);

        private const int AttachParentProcess = -1;

        public static bool IsCliInvocation(string[] args) {
            if (args == null || args.Length == 0) return false;
            foreach (string arg in args) {
                if (String.IsNullOrEmpty(arg)) continue;
                if (arg.Equals("--gui", StringComparison.OrdinalIgnoreCase)) return false;
                if (CliFlags.Contains(arg)) return true;
                int eq = arg.IndexOf('=');
                if (eq > 0 && CliFlags.Contains(arg.Substring(0, eq))) return true;
            }
            return false;
        }

        public static void EnsureConsoleAttached() {
            if (!Console.IsOutputRedirected || !Console.IsErrorRedirected) {
                AttachConsole(AttachParentProcess);
            }
        }

        // Each argument is quoted according to the Windows C runtime's argv rules.
        public static string Argument(string value) {
            var result = new StringBuilder("\"");
            int slashes = 0;
            foreach (char ch in value) {
                if (ch == '\\') { slashes++; continue; }
                if (ch == '"') result.Append('\\', slashes * 2 + 1);
                else result.Append('\\', slashes);
                slashes = 0;
                result.Append(ch);
            }
            return result.Append('\\', slashes * 2).Append('"').ToString();
        }

        public static ProcessStartInfo CreateEngineInfo(IEnumerable<string> moduleArgs, string workingDir) {
            string app = AppDomain.CurrentDomain.BaseDirectory;
            var args = new List<string> { "-I", "-B", "-X", "utf8", "-u", "-m", "tex2word" };
            args.AddRange(moduleArgs);
            var info = new ProcessStartInfo {
                FileName = Path.Combine(app, "runtime", "python.exe"),
                Arguments = String.Join(" ", args.Select(Argument)),
                WorkingDirectory = String.IsNullOrEmpty(workingDir) ? Environment.CurrentDirectory : workingDir,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };
            info.EnvironmentVariables["PYPANDOC_PANDOC"] = Path.Combine(app, "runtime", "Lib",
                "site-packages", "pypandoc", "files", "pandoc.exe");
            info.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            info.EnvironmentVariables["PYTHONUTF8"] = "1";
            return info;
        }

        public static int Run(string[] rawArgs) {
            try { Console.OutputEncoding = Encoding.UTF8; } catch { }
            var filtered = rawArgs.Where(a => !String.Equals(a, "--cli", StringComparison.OrdinalIgnoreCase)).ToList();
            var info = CreateEngineInfo(filtered, Environment.CurrentDirectory);
            using (var process = new Process { StartInfo = info }) {
                var outLock = new object();
                process.OutputDataReceived += (s, e) => {
                    if (e.Data != null) lock (outLock) Console.Out.WriteLine(e.Data);
                };
                process.ErrorDataReceived += (s, e) => {
                    if (e.Data != null) lock (outLock) Console.Error.WriteLine(e.Data);
                };
                ConsoleCancelEventHandler cancelHandler = (s, e) => {
                    e.Cancel = true;
                    try { if (!process.HasExited) process.Kill(); } catch { }
                };
                Console.CancelKeyPress += cancelHandler;
                try {
                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    process.WaitForExit();
                    return process.ExitCode;
                } finally {
                    Console.CancelKeyPress -= cancelHandler;
                }
            }
        }
    }
}
