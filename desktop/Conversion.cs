using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeX2Word {
    internal sealed class Conversion {
        public readonly string Source;
        public string Output;
        public bool CustomOutput, Replace, Complete;
        public string Error, Warning;

        public Conversion(string source) { Source = source; }

        public void ChooseDefault(string directory, HashSet<string> reserved) {
            string stem = Path.GetFileNameWithoutExtension(Source) + "-" +
                Path.GetExtension(Source).Substring(1).ToLowerInvariant() + "-MathType";
            string prefix = Path.Combine(directory ?? Path.GetDirectoryName(Source), stem);
            Output = prefix + ".docx";
            int number = 2;
            while (File.Exists(Output) || reserved.Contains(Output))
                Output = prefix + " (" + number++ + ").docx";
            Replace = false;
        }

        // Each argument is quoted according to the Windows C runtime's argv rules.
        private static string Argument(string value) {
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

        public Task Run(string language, IProgress<string> progress) {
            return Task.Run(() => {
                string app = AppDomain.CurrentDomain.BaseDirectory;
                var args = new List<string> { "-I", "-B", "-X", "utf8", "-u", "-m", "tex2word",
                    Source, "-o", Output, "--lang", language };
                if (Replace) args.Add("--force");
                var info = new ProcessStartInfo {
                    FileName = Path.Combine(app, "runtime", "python.exe"),
                    Arguments = String.Join(" ", args.Select(Argument)),
                    WorkingDirectory = app, UseShellExecute = false, CreateNoWindow = true,
                    RedirectStandardOutput = true, RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
                };
                info.EnvironmentVariables["PYPANDOC_PANDOC"] = Path.Combine(app, "runtime", "Lib",
                    "site-packages", "pypandoc", "files", "pandoc.exe");
                string lastError = null;
                var warnings = new List<string>();
                using (var process = new Process { StartInfo = info }) {
                    process.OutputDataReceived += (s, e) => {
                        if (String.IsNullOrWhiteSpace(e.Data)) return;
                        if (e.Data.StartsWith("未解析") || e.Data.StartsWith("未排版")) {
                            warnings.Add(e.Data);
                        } else if (e.Data.StartsWith("正在") || e.Data.StartsWith("识别到") ||
                                   e.Data.StartsWith("已识别") || e.Data.StartsWith("MathType 已")) {
                            progress.Report(e.Data);
                        }
                    };
                    process.ErrorDataReceived += (s, e) => {
                        if (!String.IsNullOrWhiteSpace(e.Data)) lastError = e.Data;
                    };
                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    process.WaitForExit();
                    if (process.ExitCode != 0 || !File.Exists(Output))
                        throw new IOException(lastError ?? "转换引擎未生成 Word 文件。");
                }
                Warning = String.Join(Environment.NewLine, warnings);
            });
        }
    }
}
