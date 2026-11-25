using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace TeX2Word {
    internal enum ConversionState {
        Pending,
        Running,
        Done,
        Failed
    }

    internal sealed class Conversion : INotifyPropertyChanged {
        static readonly Regex SavedCountPattern = new Regex(@"（(\d+)\s*个\s*MathType", RegexOptions.Compiled);
        static readonly Regex MathProgressPattern = new Regex(@"(\d+)/(\d+)\s*个公式", RegexOptions.Compiled);
        static readonly Regex DownloadProgressPattern = new Regex(@"正在首次下载.*?\((\d+)/(\d+)\)[：:]\s*(\d+)%", RegexOptions.Compiled);

        public readonly string Source;
        public string FileName { get; private set; }
        public string FileSizeText { get; private set; }
        public string FormatTag { get; private set; }

        string _output;
        ConversionState _state = ConversionState.Pending;
        string _liveMessage;
        string _error;
        string _warning;
        int _equationCount = -1;

        public bool CustomOutput, Replace;

        public event PropertyChangedEventHandler PropertyChanged;

        public Conversion(string source) {
            Source = source;
            FileName = Path.GetFileName(source);
            string ext = Path.GetExtension(source).ToLowerInvariant();
            if (ext == ".pdf") FormatTag = "PDF";
            else if (ext == ".tex") FormatTag = "TEX";
            else FormatTag = "MD";
            try {
                long bytes = new FileInfo(source).Length;
                FileSizeText = FormatBytes(bytes);
            } catch {
                FileSizeText = "";
            }
        }

        public string Output {
            get { return _output; }
            set {
                if (_output == value) return;
                _output = value;
                Notify("Output");
                Notify("DetailText");
                Notify("ToolTipSummary");
            }
        }

        public ConversionState State {
            get { return _state; }
            set {
                if (_state == value) return;
                _state = value;
                Notify("State");
                Notify("StateName");
                Notify("StatusText");
                Notify("DetailText");
                Notify("IsPending");
                Notify("IsRunning");
                Notify("IsDone");
                Notify("IsFailed");
                Notify("CanEditOutput");
                Notify("CanOpen");
                Notify("CanRetry");
            }
        }

        public string StateName { get { return _state.ToString(); } }
        public bool IsPending { get { return _state == ConversionState.Pending; } }
        public bool IsRunning { get { return _state == ConversionState.Running; } }
        public bool IsDone { get { return _state == ConversionState.Done; } }
        public bool IsFailed { get { return _state == ConversionState.Failed; } }
        public bool CanEditOutput { get { return _state == ConversionState.Pending || _state == ConversionState.Failed; } }
        public bool CanOpen { get { return _state == ConversionState.Done; } }
        public bool CanRetry { get { return _state == ConversionState.Failed; } }
        public bool Complete { get { return _state == ConversionState.Done; } }

        public string Error {
            get { return _error; }
            set { _error = value; Notify("Error"); Notify("DetailText"); Notify("ToolTipSummary"); }
        }

        public string Warning {
            get { return _warning; }
            set { _warning = value; Notify("Warning"); Notify("StatusText"); Notify("ToolTipSummary"); }
        }

        public string LiveMessage {
            get { return _liveMessage; }
            set { _liveMessage = value; Notify("LiveMessage"); Notify("DetailText"); }
        }

        public int EquationCount {
            get { return _equationCount; }
            set { _equationCount = value; Notify("EquationCount"); Notify("StatusText"); }
        }

        public string StatusText {
            get {
                switch (_state) {
                    case ConversionState.Running:
                        return "转换中";
                    case ConversionState.Done:
                        string suffix = _equationCount >= 0 ? String.Format(" · {0} 公式", _equationCount) : "";
                        return (String.IsNullOrEmpty(_warning) ? "已完成" : "已完成 · 有提示") + suffix;
                    case ConversionState.Failed:
                        return "转换失败";
                    default:
                        return "等待中";
                }
            }
        }

        public string DetailText {
            get {
                if (_state == ConversionState.Running && !String.IsNullOrEmpty(_liveMessage))
                    return _liveMessage;
                if (_state == ConversionState.Failed && !String.IsNullOrEmpty(_error))
                    return _error;
                return "输出至 " + (_output ?? "");
            }
        }

        public string ToolTipSummary {
            get {
                var lines = new List<string> { "源文件：" + Source, "输出至：" + (_output ?? "") };
                if (!String.IsNullOrEmpty(_warning)) lines.Add("提示：" + _warning);
                if (!String.IsNullOrEmpty(_error)) lines.Add("错误：" + _error);
                return String.Join(Environment.NewLine, lines);
            }
        }

        void Notify(string name) {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }

        static string FormatBytes(long bytes) {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.#") + " KB";
            return (bytes / (1024.0 * 1024.0)).ToString("0.##") + " MB";
        }

        public void ResetForRetry() {
            Error = null;
            Warning = null;
            LiveMessage = null;
            EquationCount = -1;
            State = ConversionState.Pending;
        }

        public void ChooseDefault(string directory, HashSet<string> reserved) {
            string ext = Path.GetExtension(Source).TrimStart('.').ToLowerInvariant();
            string stem = Path.GetFileNameWithoutExtension(Source) + "-" + ext + "-MathType";
            string prefix = Path.Combine(directory ?? Path.GetDirectoryName(Source), stem);
            string candidate = prefix + ".docx";
            int number = 2;
            while (File.Exists(candidate) || reserved.Contains(candidate))
                candidate = prefix + " (" + number++ + ").docx";
            Output = candidate;
            Replace = false;
        }

        public Task Run(string language, Action<string, double?> onProgress, Action<string> onLog, CancellationToken token) {
            return Task.Run(() => {
                string app = AppDomain.CurrentDomain.BaseDirectory;
                var args = new List<string> { Source, "-o", Output, "--lang", language };
                if (Replace) args.Add("--force");
                var info = CliRunner.CreateEngineInfo(args, app);
                string lastError = null;
                var warnings = new List<string>();
                int parsedCount = -1;

                using (var process = new Process { StartInfo = info }) {
                    process.OutputDataReceived += (s, e) => {
                        if (String.IsNullOrWhiteSpace(e.Data)) return;
                        string line = e.Data.Trim();
                        if (onLog != null) onLog(line);
                        if (line.StartsWith("未解析") || line.StartsWith("未排版")) {
                            warnings.Add(line);
                        } else if (line.StartsWith("已保存")) {
                            var m = SavedCountPattern.Match(line);
                            if (m.Success) Int32.TryParse(m.Groups[1].Value, out parsedCount);
                            if (onProgress != null) onProgress(line, 1.0);
                        } else if (line.StartsWith("MathType 已")) {
                            double? ratio = null;
                            var m = MathProgressPattern.Match(line);
                            if (m.Success) {
                                int cur, total;
                                if (Int32.TryParse(m.Groups[1].Value, out cur) &&
                                    Int32.TryParse(m.Groups[2].Value, out total) && total > 0)
                                    ratio = 0.45 + 0.50 * cur / total;
                            }
                            if (onProgress != null) onProgress(line, ratio);
                        } else if (line.StartsWith("正在首次下载")) {
                            double? ratio = null;
                            var dm = DownloadProgressPattern.Match(line);
                            if (dm.Success) {
                                int cur, total, pct;
                                if (Int32.TryParse(dm.Groups[1].Value, out cur) &&
                                    Int32.TryParse(dm.Groups[2].Value, out total) &&
                                    Int32.TryParse(dm.Groups[3].Value, out pct) && total > 0)
                                    ratio = 0.04 + 0.24 * ((cur - 1) + pct / 100.0) / total;
                            }
                            if (onProgress != null) onProgress(line, ratio);
                        } else if (line.StartsWith("正在") || line.StartsWith("识别到") || line.StartsWith("已识别")) {
                            double? ratio = line.StartsWith("正在整理") ? 0.38 : (line.StartsWith("识别到") ? 0.42 : (double?)null);
                            if (onProgress != null) onProgress(line, ratio);
                        }
                    };
                    process.ErrorDataReceived += (s, e) => {
                        if (String.IsNullOrWhiteSpace(e.Data)) return;
                        string line = e.Data.Trim();
                        lastError = line;
                        if (onLog != null) onLog("[ERR] " + line);
                    };

                    process.Start();
                    using (token.Register(() => { try { if (!process.HasExited) process.Kill(); } catch { } })) {
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();
                        process.WaitForExit();
                    }
                    token.ThrowIfCancellationRequested();
                    if (process.ExitCode != 0 || !File.Exists(Output))
                        throw new IOException(lastError ?? "转换引擎未生成 Word 文件。");
                }
                EquationCount = parsedCount;
                Warning = String.Join(Environment.NewLine, warnings);
            }, token);
        }
    }
}
