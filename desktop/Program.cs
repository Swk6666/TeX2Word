using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;

[assembly: AssemblyTitle("TeX2Word")]
[assembly: AssemblyProduct("TeX2Word")]
[assembly: AssemblyDescription("将 PDF、TeX 和 Markdown 转换为含可编辑 MathType 公式的 Word 文档")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace TeX2Word {
    internal sealed class RelayCommand : ICommand {
        readonly Action<object> _execute;
        readonly Func<object, bool> _canExecute;

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null) {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged;
        public bool CanExecute(object parameter) { return _canExecute == null || _canExecute(parameter); }
        public void Execute(object parameter) { _execute(parameter); }
        public void RaiseCanExecuteChanged() {
            var h = CanExecuteChanged;
            if (h != null) h(this, EventArgs.Empty);
        }
    }

    internal static class Program {
        [STAThread]
        static int Main(string[] args) {
            if (CliRunner.IsCliInvocation(args)) {
                CliRunner.EnsureConsoleAttached();
                return CliRunner.Run(args);
            }
            var files = args.Where(a => !String.Equals(a, "--gui", StringComparison.OrdinalIgnoreCase)).ToArray();
            var app = new Application { ShutdownMode = ShutdownMode.OnMainWindowClose };
            var controller = new MainController(app, files);
            app.Run(controller.Window);
            return 0;
        }
    }

    internal sealed class MainController {
        static readonly HashSet<string> SupportedExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            ".pdf", ".tex", ".md", ".markdown"
        };

        const string MoonPath = "M20.35,14.35 A8.5,8.5 0 0,1 9.65,3.65 A8.5,8.5 0 1,0 20.35,14.35 Z";
        const string SunPath  = "M12,7 A5,5 0 1,0 12,17 A5,5 0 1,0 12,7 Z M12,2 V4 M12,20 V22 M4.9,4.9 L6.3,6.3 M17.7,17.7 L19.1,19.1 M2,12 H4 M20,12 H22 M4.9,19.1 L6.3,17.7 M17.7,6.3 L19.1,4.9";

        readonly Application _app;
        public readonly Window Window;
        public readonly ObservableCollection<Conversion> Jobs = new ObservableCollection<Conversion>();

        public RelayCommand OpenItemCommand { get; private set; }
        public RelayCommand RevealItemCommand { get; private set; }
        public RelayCommand RenameItemCommand { get; private set; }
        public RelayCommand RetryItemCommand { get; private set; }
        public RelayCommand InspectErrorCommand { get; private set; }
        public RelayCommand RemoveItemCommand { get; private set; }

        readonly Border _emptyStateCard, _queueCard, _dragOverlay, _logDrawer;
        readonly ItemsControl _jobsList;
        readonly TextBlock _queueCountText, _outputDirText, _statusText, _envText, _toggleLogLabel, _startButtonLabel;
        readonly Ellipse _envDot;
        readonly System.Windows.Shapes.Path _themeIcon;
        readonly Button _envButton, _cliButton, _themeButton, _emptyPickButton, _addMoreButton;
        readonly Button _clearDoneButton, _clearAllButton, _pickDirButton, _resetDirButton;
        readonly Button _toggleLogButton, _copyLogButton, _clearLogButton, _startButton, _stopButton;
        readonly RadioButton _langZhRadio, _langEnRadio;
        readonly ProgressBar _progressBar;
        readonly TextBox _logTextBox;

        string _outputDirectory;
        bool _busy;
        bool _darkMode;
        bool _hasWord, _hasMathType;
        CancellationTokenSource _cts;

        public MainController(Application app, string[] initialFiles) {
            _app = app;
            _darkMode = ShellDialogs.SystemPrefersDarkTheme();

            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("TeX2Word.MainWindow.xaml")) {
                if (stream == null) throw new InvalidOperationException("缺少内嵌界面资源 TeX2Word.MainWindow.xaml。");
                Window = (Window)XamlReader.Load(stream);
            }

            ApplyWindowIcon();
            Theme.Apply(_app.Resources, Window, _darkMode);

            OpenItemCommand     = new RelayCommand(p => OpenJob(p as Conversion));
            RevealItemCommand   = new RelayCommand(p => { var j = p as Conversion; if (j != null) ShellDialogs.RevealInExplorer(j.Output); });
            RenameItemCommand   = new RelayCommand(p => RenameJob(p as Conversion), p => !_busy);
            RetryItemCommand    = new RelayCommand(p => RetryJob(p as Conversion), p => !_busy);
            InspectErrorCommand = new RelayCommand(p => InspectError(p as Conversion));
            RemoveItemCommand   = new RelayCommand(p => RemoveJob(p as Conversion), p => !_busy);

            Window.DataContext = this;

            _emptyStateCard   = (Border)Window.FindName("EmptyStateCard");
            _queueCard        = (Border)Window.FindName("QueueCard");
            _dragOverlay      = (Border)Window.FindName("DragOverlay");
            _logDrawer        = (Border)Window.FindName("LogDrawer");
            _jobsList         = (ItemsControl)Window.FindName("JobsList");
            _queueCountText   = (TextBlock)Window.FindName("QueueCountText");
            _outputDirText    = (TextBlock)Window.FindName("OutputDirText");
            _statusText       = (TextBlock)Window.FindName("StatusText");
            _envText          = (TextBlock)Window.FindName("EnvText");
            _envDot           = (Ellipse)Window.FindName("EnvDot");
            _themeIcon        = (System.Windows.Shapes.Path)Window.FindName("ThemeIcon");
            _toggleLogLabel   = (TextBlock)Window.FindName("ToggleLogLabel");
            _startButtonLabel = (TextBlock)Window.FindName("StartButtonLabel");
            _envButton        = (Button)Window.FindName("EnvButton");
            _cliButton        = (Button)Window.FindName("CliButton");
            _themeButton      = (Button)Window.FindName("ThemeButton");
            _emptyPickButton  = (Button)Window.FindName("EmptyPickButton");
            _addMoreButton    = (Button)Window.FindName("AddMoreButton");
            _clearDoneButton  = (Button)Window.FindName("ClearDoneButton");
            _clearAllButton   = (Button)Window.FindName("ClearAllButton");
            _pickDirButton    = (Button)Window.FindName("PickDirButton");
            _resetDirButton   = (Button)Window.FindName("ResetDirButton");
            _toggleLogButton  = (Button)Window.FindName("ToggleLogButton");
            _copyLogButton    = (Button)Window.FindName("CopyLogButton");
            _clearLogButton   = (Button)Window.FindName("ClearLogButton");
            _startButton      = (Button)Window.FindName("StartButton");
            _stopButton       = (Button)Window.FindName("StopButton");
            _langZhRadio      = (RadioButton)Window.FindName("LangZhRadio");
            _langEnRadio      = (RadioButton)Window.FindName("LangEnRadio");
            _progressBar      = (ProgressBar)Window.FindName("MainProgressBar");
            _logTextBox       = (TextBox)Window.FindName("LogTextBox");

            _jobsList.ItemsSource = Jobs;
            UpdateThemeIcon();

            Window.SourceInitialized += (s, e) => Theme.Apply(_app.Resources, Window, _darkMode);
            Window.PreviewDragEnter += OnDragOver;
            Window.PreviewDragOver  += OnDragOver;
            Window.PreviewDragLeave += (s, e) => {
                var pos = e.GetPosition(Window);
                if (pos.X <= 0 || pos.Y <= 0 || pos.X >= Window.ActualWidth || pos.Y >= Window.ActualHeight)
                    _dragOverlay.Visibility = Visibility.Collapsed;
            };
            Window.PreviewDrop += OnDrop;
            Window.PreviewKeyDown += OnKeyDown;
            Window.Closing += (s, e) => {
                if (_busy) {
                    e.Cancel = true;
                    MessageBox.Show(Window, "正在进行转换，请等待当前队列完成或点击“停止队列”后再关闭。",
                        "正在转换", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            };

            _envButton.Click       += (s, e) => ShowEnvDialog();
            _cliButton.Click       += (s, e) => ShowCliDialog();
            _themeButton.Click     += (s, e) => ToggleTheme();
            _emptyPickButton.Click += (s, e) => PickFiles();
            var emptyDrop = Window.FindName("EmptyDropSurface") as UIElement;
            if (emptyDrop != null) emptyDrop.MouseLeftButtonUp += (s, e) => {
                if (!_busy && !(e.OriginalSource is Button)) PickFiles();
            };
            _addMoreButton.Click   += (s, e) => PickFiles();
            _clearDoneButton.Click += (s, e) => ClearDoneJobs();
            _clearAllButton.Click  += (s, e) => ClearAllJobs();
            _pickDirButton.Click   += (s, e) => PickOutputDirectory();
            _resetDirButton.Click  += (s, e) => { _outputDirectory = null; RefreshOutputs(); };
            _toggleLogButton.Click += (s, e) => ToggleLogDrawer();
            _copyLogButton.Click   += (s, e) => {
                if (!String.IsNullOrEmpty(_logTextBox.Text)) {
                    try { Clipboard.SetText(_logTextBox.Text); } catch { }
                    _statusText.Text = "已复制日志到剪贴板。";
                }
            };
            _clearLogButton.Click += (s, e) => _logTextBox.Clear();
            _startButton.Click    += (s, e) => ConvertAll();
            _stopButton.Click     += (s, e) => {
                if (_cts != null && !_cts.IsCancellationRequested) {
                    _cts.Cancel();
                    _stopButton.IsEnabled = false;
                    _statusText.Text = "正在停止当前转换任务…";
                }
            };

            RefreshEnvironmentStatus();
            AddFiles(initialFiles);
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern IntPtr ExtractAssociatedIcon(IntPtr hInst, string lpIconPath, out ushort lpiIcon);

        [DllImport("user32.dll")]
        static extern bool DestroyIcon(IntPtr hIcon);

        void ApplyWindowIcon() {
            try {
                ushort idx;
                string exe = Assembly.GetExecutingAssembly().Location;
                IntPtr hIcon = ExtractAssociatedIcon(IntPtr.Zero, exe, out idx);
                if (hIcon != IntPtr.Zero) {
                    try {
                        Window.Icon = Imaging.CreateBitmapSourceFromHIcon(
                            hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    } finally {
                        DestroyIcon(hIcon);
                    }
                }
            } catch { }
        }

        void RefreshEnvironmentStatus() {
            _hasWord = ShellDialogs.HasWord();
            _hasMathType = ShellDialogs.HasMathType();
            if (_hasWord && _hasMathType) {
                _envText.Text = "Word · MathType 就绪";
                _envDot.SetResourceReference(Shape.FillProperty, "SuccessDot");
            } else if (!_hasWord && !_hasMathType) {
                _envText.Text = "缺少 Word 与 MathType";
                _envDot.SetResourceReference(Shape.FillProperty, "WarnDot");
            } else if (!_hasWord) {
                _envText.Text = "未检测到 Word";
                _envDot.SetResourceReference(Shape.FillProperty, "WarnDot");
            } else {
                _envText.Text = "未检测到 MathType 7";
                _envDot.SetResourceReference(Shape.FillProperty, "WarnDot");
            }
        }

        void ShowEnvDialog() {
            RefreshEnvironmentStatus();
            string msg = String.Format(
                "运行环境检测结果（Windows 桌面版）：\n\n" +
                "• Microsoft Word 桌面版：{0}\n" +
                "• MathType 7 桌面版 (Equation.DSMT4)：{1}\n\n" +
                "说明：TeX2Word 通过 Word COM 与 MathType 桌面 SDK 写入原生可编辑 MathType 公式对象。\n" +
                "如自定义了 MathType 安装目录，可设置环境变量 MATHTYPE_DLL 指向 64 位 MT6.dll。",
                _hasWord ? "已检测到 ✓" : "未检测到 ✗",
                _hasMathType ? "已检测到 ✓" : "未检测到 ✗");
            MessageBox.Show(Window, msg, "运行环境检测", MessageBoxButton.OK,
                (_hasWord && _hasMathType) ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        void ToggleTheme() {
            _darkMode = !_darkMode;
            Theme.Apply(_app.Resources, Window, _darkMode);
            UpdateThemeIcon();
        }

        void UpdateThemeIcon() {
            _themeIcon.Data = Geometry.Parse(_darkMode ? SunPath : MoonPath);
        }

        void ToggleLogDrawer() {
            bool show = _logDrawer.Visibility != Visibility.Visible;
            _logDrawer.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            _toggleLogLabel.Text = show ? "隐藏日志" : "日志";
        }

        void AppendLog(string line) {
            Window.Dispatcher.BeginInvoke(new Action(() => {
                _logTextBox.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + line + Environment.NewLine);
                _logTextBox.ScrollToEnd();
            }));
        }

        void OnDragOver(object sender, DragEventArgs e) {
            if (!_busy && e.Data.GetDataPresent(DataFormats.FileDrop)) {
                e.Effects = DragDropEffects.Copy;
                _dragOverlay.Visibility = Visibility.Visible;
            } else {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        void OnDrop(object sender, DragEventArgs e) {
            _dragOverlay.Visibility = Visibility.Collapsed;
            if (!_busy && e.Data.GetDataPresent(DataFormats.FileDrop)) {
                var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (paths != null) AddFiles(paths);
            }
            e.Handled = true;
        }

        void OnKeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.L && Keyboard.Modifiers == ModifierKeys.Control) {
                ToggleLogDrawer();
                e.Handled = true;
                return;
            }
            if (_busy) return;
            if (e.Key == Key.O && Keyboard.Modifiers == ModifierKeys.Control) {
                PickFiles();
                e.Handled = true;
            } else if (e.Key == Key.F5 && _startButton.IsEnabled) {
                ConvertAll();
                e.Handled = true;
            }
        }

        void PickFiles() {
            var dialog = new OpenFileDialog {
                Multiselect = true,
                Title = "添加 PDF、TeX 或 Markdown 文档",
                Filter = "支持的文档 (*.pdf;*.tex;*.md)|*.pdf;*.tex;*.md;*.markdown|PDF 文档 (*.pdf)|*.pdf|TeX 主文件 (*.tex)|*.tex|Markdown 文档 (*.md)|*.md;*.markdown"
            };
            if (dialog.ShowDialog(Window) == true) AddFiles(dialog.FileNames);
        }

        void PickOutputDirectory() {
            string chosen = ShellDialogs.PickFolder(Window, "选择 Word 文档输出目录", _outputDirectory);
            if (!String.IsNullOrEmpty(chosen)) {
                _outputDirectory = chosen;
                RefreshOutputs();
            }
        }

        void AddFiles(IEnumerable<string> paths) {
            if (paths == null) return;
            var known = new HashSet<string>(Jobs.Select(j => j.Source), StringComparer.OrdinalIgnoreCase);
            int added = 0, skipped = 0;
            foreach (string candidate in paths) {
                if (String.IsNullOrWhiteSpace(candidate)) continue;
                string ext = System.IO.Path.GetExtension(candidate);
                if (!SupportedExt.Contains(ext) || !File.Exists(candidate)) { skipped++; continue; }
                string full = System.IO.Path.GetFullPath(candidate);
                if (!known.Add(full)) continue;
                Jobs.Add(new Conversion(full));
                added++;
            }
            if (added > 0 || skipped > 0) {
                _statusText.Text = added > 0
                    ? String.Format("已加入 {0} 份文档，当前队列共 {1} 份。", added, Jobs.Count)
                    : String.Format("当前队列共 {0} 份文档。", Jobs.Count);
                if (skipped > 0) _statusText.Text += String.Format(" 已跳过 {0} 个不支持或无效的路径。", skipped);
            }
            RefreshOutputs();
        }

        void RefreshOutputs() {
            _outputDirText.Text = String.IsNullOrEmpty(_outputDirectory) ? "默认：各输入文件所在的目录" : _outputDirectory;
            _outputDirText.ToolTip = _outputDirText.Text;
            _resetDirButton.Visibility = String.IsNullOrEmpty(_outputDirectory) ? Visibility.Collapsed : Visibility.Visible;

            var reserved = new HashSet<string>(
                Jobs.Where(j => j.Complete || j.CustomOutput).Select(j => j.Output).Where(p => !String.IsNullOrEmpty(p)),
                StringComparer.OrdinalIgnoreCase);
            foreach (var job in Jobs) {
                if (!job.Complete && !job.CustomOutput) job.ChooseDefault(_outputDirectory, reserved);
                if (!String.IsNullOrEmpty(job.Output)) reserved.Add(job.Output);
            }
            SyncUiState();
        }

        void SyncUiState() {
            bool hasItems = Jobs.Count > 0;
            _emptyStateCard.Visibility = hasItems ? Visibility.Collapsed : Visibility.Visible;
            _queueCard.Visibility = hasItems ? Visibility.Visible : Visibility.Collapsed;

            int doneCount = Jobs.Count(j => j.IsDone);
            int pendingCount = Jobs.Count(j => !j.IsDone);
            _queueCountText.Text = doneCount > 0
                ? String.Format("{0} 份文件 · 已完成 {1}", Jobs.Count, doneCount)
                : String.Format("{0} 份文件", Jobs.Count);

            _addMoreButton.IsEnabled   = !_busy;
            _clearDoneButton.IsEnabled = !_busy && doneCount > 0;
            _clearAllButton.IsEnabled  = !_busy && hasItems;
            _pickDirButton.IsEnabled   = !_busy;
            _resetDirButton.IsEnabled  = !_busy && !String.IsNullOrEmpty(_outputDirectory);
            _langZhRadio.IsEnabled     = !_busy;
            _langEnRadio.IsEnabled     = !_busy;
            _startButton.IsEnabled     = !_busy && pendingCount > 0;
            _stopButton.Visibility     = _busy ? Visibility.Visible : Visibility.Collapsed;
            _stopButton.IsEnabled      = _busy;
            _startButtonLabel.Text     = _busy ? "转换中…" : (pendingCount > 0 && doneCount > 0 ? "继续转换" : "开始转换");

            RenameItemCommand.RaiseCanExecuteChanged();
            RetryItemCommand.RaiseCanExecuteChanged();
            RemoveItemCommand.RaiseCanExecuteChanged();
        }

        void OpenJob(Conversion job) {
            if (job == null) return;
            if (job.IsFailed) { InspectError(job); return; }
            if (!job.IsDone) return;
            try {
                Process.Start(new ProcessStartInfo(job.Output) { UseShellExecute = true });
            } catch (Exception ex) {
                MessageBox.Show(Window, ex.Message, "打开 Word 文档", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void InspectError(Conversion job) {
            if (job == null) return;
            string detail = !String.IsNullOrEmpty(job.Error) ? job.Error : "未知错误，请打开底部“日志”面板查看完整输出。";
            MessageBox.Show(Window, detail, "转换失败详情 · " + job.FileName, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        void RenameJob(Conversion job) {
            if (job == null || _busy || job.IsDone) return;
            var dialog = new SaveFileDialog {
                Filter = "Word 文档 (*.docx)|*.docx",
                DefaultExt = "docx",
                OverwritePrompt = true,
                FileName = System.IO.Path.GetFileName(job.Output),
                InitialDirectory = System.IO.Path.GetDirectoryName(job.Output)
            };
            if (dialog.ShowDialog(Window) != true) return;
            string path = System.IO.Path.GetFullPath(dialog.FileName);
            if (Jobs.Any(x => x != job && String.Equals(x.Output, path, StringComparison.OrdinalIgnoreCase))) {
                MessageBox.Show(Window, "队列中已有其他任务使用这个输出路径，请更换文件名。",
                    "输出路径冲突", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            job.Output = path;
            job.CustomOutput = true;
            job.Replace = File.Exists(path);
            RefreshOutputs();
        }

        void RetryJob(Conversion job) {
            if (job == null || _busy) return;
            job.ResetForRetry();
            SyncUiState();
            _statusText.Text = "已重置 " + job.FileName + "，点击“开始转换”即可重试。";
        }

        void RemoveJob(Conversion job) {
            if (job == null || _busy) return;
            Jobs.Remove(job);
            RefreshOutputs();
            if (Jobs.Count == 0) {
                _progressBar.Value = 0;
                _statusText.Text = "准备就绪 · 添加文件后点击“开始转换”（或按 F5）";
            }
        }

        void ClearDoneJobs() {
            if (_busy) return;
            foreach (var j in Jobs.Where(x => x.IsDone).ToArray()) Jobs.Remove(j);
            RefreshOutputs();
        }

        void ClearAllJobs() {
            if (_busy) return;
            Jobs.Clear();
            _progressBar.Value = 0;
            _statusText.Text = "准备就绪 · 添加文件后点击“开始转换”（或按 F5）";
            RefreshOutputs();
        }

        void ShowCliDialog() {
            var dlg = new Window {
                Owner = Window,
                Title = "命令行（CLI）快速参考",
                Width = 620,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = (Brush)_app.Resources["CanvasBg"],
                Foreground = (Brush)_app.Resources["TextPrimary"],
                FontFamily = Window.FontFamily,
                UseLayoutRounding = true
            };
            dlg.SourceInitialized += (s, e) => Theme.Apply(_app.Resources, dlg, _darkMode);

            var root = new StackPanel { Margin = new Thickness(20) };
            root.Children.Add(new TextBlock {
                Text = "安装版已自动将 tex2word 加入用户 PATH，可在 PowerShell 或 CMD 中直接运行：",
                FontSize = 13, TextWrapping = TextWrapping.Wrap,
                Foreground = (Brush)_app.Resources["TextSecondary"],
                Margin = new Thickness(0, 0, 0, 12)
            });

            string cmds =
                "# 转换 PDF / TeX / Markdown 为含 MathType 公式的 Word\r\n" +
                "tex2word \"paper.pdf\" -o \"paper-MathType.docx\"\r\n" +
                "tex2word \"main.tex\" --force\r\n" +
                "tex2word \"notes.md\" -o \"notes-MathType.docx\"\r\n\r\n" +
                "# 混合批量转换到指定目录\r\n" +
                "tex2word \"a.pdf\" \"b.tex\" \"c.md\" -d \"D:\\Output\"\r\n\r\n" +
                "# 校验成品文档中的 MathType 原生对象\r\n" +
                "tex2word --verify \"paper-MathType.docx\"";

            var box = new TextBox {
                Text = cmds, IsReadOnly = true, AcceptsReturn = true,
                FontFamily = new FontFamily("Consolas, Microsoft YaHei UI"),
                FontSize = 12, Padding = new Thickness(12),
                Background = (Brush)_app.Resources["LogBg"],
                Foreground = (Brush)_app.Resources["LogFg"],
                BorderBrush = (Brush)_app.Resources["LogBorder"],
                BorderThickness = new Thickness(1)
            };
            root.Children.Add(box);

            var footer = new StackPanel {
                Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            var copyBtn = new Button {
                Content = "复制命令", Style = (Style)Window.Resources["SecondaryButtonStyle"],
                Margin = new Thickness(0, 0, 8, 0)
            };
            copyBtn.Click += (s, e) => { try { Clipboard.SetText(cmds); copyBtn.Content = "已复制 ✓"; } catch { } };
            var closeBtn = new Button {
                Content = "关闭", Style = (Style)Window.Resources["PrimaryButtonStyle"],
                Padding = new Thickness(18, 7, 18, 7), IsDefault = true, IsCancel = true
            };
            closeBtn.Click += (s, e) => dlg.Close();
            footer.Children.Add(copyBtn);
            footer.Children.Add(closeBtn);
            root.Children.Add(footer);
            dlg.Content = root;
            dlg.ShowDialog();
        }

        async void ConvertAll() {
            if (_busy) return;
            var queue = Jobs.Where(j => !j.IsDone).ToArray();
            if (queue.Length == 0) return;

            RefreshEnvironmentStatus();
            if (!_hasWord) {
                MessageBox.Show(Window, "请先安装 Microsoft Word 桌面版，然后重新开始转换。",
                    "需要 Microsoft Word", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!_hasMathType) {
                MessageBox.Show(Window, "未检测到 MathType 7 桌面版（Equation.DSMT4）。请安装桌面版 MathType 7 后重试。",
                    "需要 MathType 7", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            RefreshOutputs();
            _busy = true;
            _cts = new CancellationTokenSource();
            SyncUiState();

            string lang = _langEnRadio.IsChecked == true ? "en" : "ch";
            _progressBar.IsIndeterminate = false;
            _progressBar.Value = 0;

            int succeeded = 0, failed = 0;
            bool cancelled = false;

            try {
                for (int index = 0; index < queue.Length; index++) {
                    if (_cts.IsCancellationRequested) { cancelled = true; break; }
                    var job = queue[index];
                    string prefix = String.Format("({0}/{1}) {2} · ", index + 1, queue.Length, job.FileName);
                    job.Error = null;
                    job.Warning = null;
                    job.LiveMessage = "正在启动转换引擎…";
                    job.State = ConversionState.Running;
                    _statusText.Text = prefix + "正在读取文档…";
                    AppendLog("=== 开始转换: " + job.Source + " ===");

                    double basePct = index * 100.0 / queue.Length;
                    double spanPct = 100.0 / queue.Length;
                    _progressBar.Value = Math.Max(2, basePct + spanPct * 0.08);

                    try {
                        await job.Run(lang, (msg, ratio) => {
                            Window.Dispatcher.BeginInvoke(new Action(() => {
                                if (job.State != ConversionState.Running) return;
                                job.LiveMessage = msg;
                                _statusText.Text = prefix + msg;
                                if (ratio.HasValue)
                                    _progressBar.Value = basePct + spanPct * Math.Min(1.0, Math.Max(0.05, ratio.Value));
                            }));
                        }, AppendLog, _cts.Token);

                        job.LiveMessage = null;
                        job.State = ConversionState.Done;
                        succeeded++;
                        _progressBar.Value = (index + 1) * 100.0 / queue.Length;
                    } catch (OperationCanceledException) {
                        job.ResetForRetry();
                        cancelled = true;
                        AppendLog("用户取消了当前队列。");
                        break;
                    } catch (Exception ex) {
                        failed++;
                        job.Error = ex.Message;
                        job.LiveMessage = null;
                        job.State = ConversionState.Failed;
                        AppendLog("转换失败: " + ex.Message);
                        _progressBar.Value = (index + 1) * 100.0 / queue.Length;
                    }
                    SyncUiState();
                }

                if (cancelled) {
                    _statusText.Text = String.Format("已停止队列 · 本次已完成 {0} 份，未完成项保留在队列中。", succeeded);
                } else if (failed > 0) {
                    _statusText.Text = String.Format("已完成 {0} 份，失败 {1} 份 · 可点击失败项查看详情或打开底部日志。", succeeded, failed);
                } else {
                    _statusText.Text = String.Format("全部 {0} 份文档转换完成 · 点击“打开 Word”即可直接编辑公式。", succeeded);
                }
            } finally {
                _busy = false;
                if (_cts != null) { _cts.Dispose(); _cts = null; }
                SyncUiState();
            }
        }
    }
}
