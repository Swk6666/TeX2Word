using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("TeX2Word")]
[assembly: AssemblyProduct("TeX2Word")]
[assembly: AssemblyDescription("将 TeX 和 PDF 转换为含可编辑 MathType 公式的 Word 文档")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace TeX2Word {
    internal static class Program {
        [STAThread]
        static void Main(string[] files) {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainWindow(files));
        }
    }

    internal sealed class MainWindow : Form {
        readonly ListView list = new ListView();
        readonly Button add = ButtonOf("添加文件…"), remove = ButtonOf("移除"),
            rename = ButtonOf("另存为…"), open = ButtonOf("打开结果"), start = ButtonOf("开始转换");
        readonly Button browse = ButtonOf("选择目录…"), reset = ButtonOf("使用原目录");
        readonly ComboBox language = new ComboBox();
        readonly TextBox folder = new TextBox();
        readonly Label status = new Label();
        readonly ProgressBar progress = new ProgressBar();
        string directory;
        bool busy;

        static Button ButtonOf(string caption) {
            return new Button { Text = caption, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(9, 3, 9, 3), UseVisualStyleBackColor = true,
                Anchor = AnchorStyles.Left, Margin = new Padding(0, 3, 6, 3) };
        }

        static Label Caption(string text) {
            return new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 3, 10, 3) };
        }

        public MainWindow(string[] initial) {
            SuspendLayout();
            Text = "TeX2Word · 公式文档转换";
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.White;
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(840, 510);
            MinimumSize = new Size(700, 450);
            StartPosition = FormStartPosition.CenterScreen;
            AllowDrop = true;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 7,
                Padding = new Padding(20, 14, 20, 14) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 7; i++) layout.RowStyles.Add(new RowStyle(
                i == 2 ? SizeType.Percent : SizeType.AutoSize, i == 2 ? 100 : 0));
            layout.RowStyles[6] = new RowStyle(SizeType.Absolute, 45);
            Controls.Add(layout);
            var heading = Caption("将 TeX 和 PDF 转换为 Word");
            heading.Font = new Font(Font.FontFamily, 15F, FontStyle.Bold);
            heading.Margin = new Padding(0, 0, 0, 6);
            layout.Controls.Add(heading, 0, 0);
            var toolbar = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill,
                WrapContents = false, Margin = new Padding(0, 0, 0, 4) };
            toolbar.Controls.AddRange(new Control[] { add, remove, rename, open });
            layout.Controls.Add(toolbar, 0, 1);

            list.Dock = DockStyle.Fill; list.View = View.Details; list.FullRowSelect = true;
            list.HideSelection = false; list.MultiSelect = true; list.ShowItemToolTips = true;
            list.AccessibleName = "转换文件列表"; list.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            list.Columns.Add("输入文件", 220); list.Columns.Add("Word 输出文件", 420); list.Columns.Add("状态", 108);
            list.Margin = new Padding(0, 0, 0, 10);
            layout.Controls.Add(list, 0, 2);
            list.Resize += (s, e) => {
                int width = Math.Max(400, list.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 6);
                int statusWidth = TextRenderer.MeasureText("失败，双击查看　", list.Font).Width + 12;
                list.Columns[0].Width = width * 31 / 100;
                list.Columns[1].Width = width - list.Columns[0].Width - statusWidth;
                list.Columns[2].Width = statusWidth;
            };
            var pathRow = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 4, Margin = Padding.Empty };
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pathRow.Controls.Add(Caption("输出目录"), 0, 0);
            folder.ReadOnly = true; folder.Dock = DockStyle.Fill; folder.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            folder.AccessibleName = "输出目录"; folder.Margin = new Padding(0, 2, 8, 2);
            pathRow.Controls.Add(folder, 1, 0); pathRow.Controls.Add(browse, 2, 0); pathRow.Controls.Add(reset, 3, 0);
            layout.Controls.Add(pathRow, 0, 3);
            var options = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = false, Margin = new Padding(0, 2, 0, 6) };
            options.Controls.Add(Caption("PDF 识别语言"));
            language.DropDownStyle = ComboBoxStyle.DropDownList; language.Width = 145;
            language.AccessibleName = "PDF 识别语言"; language.Items.AddRange(new object[] { "中文或中英混排", "英文" }); language.SelectedIndex = 0;
            options.Controls.Add(language);
            layout.Controls.Add(options, 0, 4);
            progress.Dock = DockStyle.Fill; progress.Height = 6; progress.Margin = new Padding(0, 4, 0, 5);
            layout.Controls.Add(progress, 0, 5);
            var actions = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            actions.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            status.Text = "可添加或拖入多个文件，公式可在桌面版 MathType 中编辑。";
            status.AutoEllipsis = true; status.Dock = DockStyle.Fill; status.TextAlign = ContentAlignment.MiddleLeft;
            status.AccessibleName = "转换进度"; actions.Controls.Add(status, 0, 0);
            start.BackColor = Color.FromArgb(25, 83, 147); start.ForeColor = Color.White;
            start.FlatStyle = FlatStyle.Flat; start.FlatAppearance.BorderSize = 0; start.Padding = new Padding(15, 5, 15, 5);
            actions.Controls.Add(start, 1, 0); layout.Controls.Add(actions, 0, 6);
            add.Click += (s, e) => PickFiles();
            remove.Click += (s, e) => { foreach (var item in Selected()) list.Items.Remove(item); RefreshOutputs(); };
            rename.Click += (s, e) => RenameOutput();
            open.Click += (s, e) => OpenSelected();
            list.DoubleClick += (s, e) => OpenSelected();
            list.SelectedIndexChanged += (s, e) => SyncButtons();
            browse.Click += (s, e) => {
                using (var dialog = new FolderBrowserDialog { Description = "选择生成 Word 文档的目录", SelectedPath = directory ?? "" }) {
                    if (dialog.ShowDialog(this) == DialogResult.OK) { directory = dialog.SelectedPath; RefreshOutputs(); }
                }
            };
            reset.Click += (s, e) => { directory = null; RefreshOutputs(); };
            start.Click += (s, e) => ConvertAll();
            DragEnter += (s, e) => e.Effect = !busy && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            DragDrop += (s, e) => { if (!busy) AddFiles((string[])e.Data.GetData(DataFormats.FileDrop)); };
            KeyPreview = true;
            KeyDown += (s, e) => {
                if (busy) return;
                if (e.Control && e.KeyCode == Keys.O) { e.Handled = true; PickFiles(); }
                if (e.KeyCode == Keys.F5 && start.Enabled) { e.Handled = true; ConvertAll(); }
            };
            FormClosing += (s, e) => {
                if (busy && e.CloseReason == CloseReason.UserClosing) {
                    e.Cancel = true;
                    MessageBox.Show(this, "正在写入 Word。请在队列完成后关闭窗口。", "正在转换");
                }
            };
            AddFiles(initial);
            ResumeLayout(true);
        }

        ListViewItem[] Rows() { return list.Items.Cast<ListViewItem>().ToArray(); }
        ListViewItem[] Selected() { return list.SelectedItems.Cast<ListViewItem>().ToArray(); }
        static Conversion Job(ListViewItem row) { return (Conversion)row.Tag; }

        void PickFiles() {
            using (var dialog = new OpenFileDialog { Multiselect = true, Title = "添加 TeX 或 PDF 文档",
                Filter = "TeX 和 PDF 文档|*.tex;*.pdf|TeX 主文件|*.tex|PDF 文档|*.pdf" })
                if (dialog.ShowDialog(this) == DialogResult.OK) AddFiles(dialog.FileNames);
        }

        void AddFiles(IEnumerable<string> paths) {
            var known = new HashSet<string>(Rows().Select(row => Job(row).Source), StringComparer.OrdinalIgnoreCase);
            int skipped = 0;
            foreach (string candidate in paths) {
                string ext = Path.GetExtension(candidate).ToLowerInvariant();
                if ((ext != ".tex" && ext != ".pdf") || !File.Exists(candidate)) { skipped++; continue; }
                string path = Path.GetFullPath(candidate);
                if (!known.Add(path)) continue;
                var row = new ListViewItem(new string[] { Path.GetFileName(path), "", "等待转换" });
                row.Tag = new Conversion(path); row.ToolTipText = path; list.Items.Add(row);
            }
            if (list.Items.Count > 0) status.Text = String.Format("列表中有 {0} 份文件。", list.Items.Count);
            if (skipped > 0) status.Text += String.Format(" 已跳过 {0} 个无效或不支持的输入。", skipped);
            RefreshOutputs();
        }

        void RefreshOutputs() {
            folder.Text = directory ?? "各输入文件所在的目录";
            var reserved = new HashSet<string>(Rows().Select(Job).Where(x => x.Complete || x.CustomOutput)
                .Select(x => x.Output), StringComparer.OrdinalIgnoreCase);
            foreach (var row in Rows()) {
                var job = Job(row);
                if (!job.Complete && !job.CustomOutput) job.ChooseDefault(directory, reserved);
                reserved.Add(job.Output);
                row.SubItems[1].Text = job.Output;
            }
            SyncButtons();
        }

        void SyncButtons() {
            var sel = Selected();
            add.Enabled = browse.Enabled = language.Enabled = !busy; reset.Enabled = !busy && directory != null;
            remove.Enabled = !busy && sel.Length > 0;
            rename.Enabled = !busy && sel.Length == 1 && !Job(sel[0]).Complete;
            open.Enabled = sel.Length == 1 && Job(sel[0]).Complete;
            start.Enabled = !busy && Rows().Any(row => !Job(row).Complete);
            start.BackColor = start.Enabled ? Color.FromArgb(25, 83, 147) : Color.FromArgb(226, 231, 237);
        }

        void RenameOutput() {
            var job = Job(Selected()[0]);
            using (var dialog = new SaveFileDialog { Filter = "Word 文档|*.docx", DefaultExt = "docx",
                OverwritePrompt = true, FileName = Path.GetFileName(job.Output), InitialDirectory = Path.GetDirectoryName(job.Output) }) {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                string path = Path.GetFullPath(dialog.FileName);
                if (Rows().Select(Job).Any(x => x != job && String.Equals(x.Output, path, StringComparison.OrdinalIgnoreCase))) {
                    MessageBox.Show(this, "队列中已有文件使用这个输出位置，请换一个文件名。", "选择输出"); return;
                }
                job.Output = path; job.CustomOutput = true; job.Replace = File.Exists(path); RefreshOutputs();
            }
        }

        void OpenSelected() {
            var row = Selected().FirstOrDefault();
            if (row == null) return;
            var job = Job(row);
            if (!String.IsNullOrEmpty(job.Error)) { MessageBox.Show(this, job.Error, "转换详情"); return; }
            if (!job.Complete) return;
            try { Process.Start(new ProcessStartInfo(job.Output) { UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "打开 Word 文档"); }
        }

        async void ConvertAll() {
            if (busy) return;
            var queue = Rows().Where(row => !Job(row).Complete).ToArray();
            if (queue.Length == 0) return;
            if (Type.GetTypeFromProgID("Word.Application") == null) {
                MessageBox.Show(this, "请先安装 Microsoft Word 桌面版，然后重新开始转换。", "需要 Word"); return;
            }
            RefreshOutputs();
            busy = true; SyncButtons();
            string lang = language.SelectedIndex == 1 ? "en" : "ch";
            progress.Value = 0; progress.Maximum = queue.Length;
            int failed = 0;
            try {
                for (int index = 0; index < queue.Length; index++) {
                    var row = queue[index]; var job = Job(row);
                    string prefix = String.Format("({0}/{1}) {2}   ", index + 1, queue.Length, row.Text);
                    row.SubItems[2].Text = "转换中"; job.Error = job.Warning = null; row.EnsureVisible();
                    status.Text = prefix + "正在读取文档…";
                    var reporter = new Progress<string>(s => {
                        if (row.SubItems[2].Text == "转换中") status.Text = prefix + s;
                    });
                    try {
                        await job.Run(lang, reporter);
                        job.Complete = true;
                        row.SubItems[2].Text = String.IsNullOrEmpty(job.Warning) ? "已完成" : "已完成，有提示";
                        row.ToolTipText = job.Source + Environment.NewLine + job.Output + Environment.NewLine + job.Warning;
                    } catch (Exception ex) {
                        failed++; job.Error = ex.Message;
                        row.SubItems[2].Text = "失败，双击查看"; row.ToolTipText = job.Error;
                    }
                    progress.Value = index + 1;
                }
                status.Text = String.Format("已完成 {0} 份，失败 {1} 份。", queue.Length - failed, failed)
                    + (failed > 0 ? "双击失败项查看原因，可再次开始转换。" : "双击已完成的文件即可打开 Word。");
            } finally { busy = false; SyncButtons(); }
        }
    }
}
