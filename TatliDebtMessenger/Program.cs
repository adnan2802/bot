using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace TatliDebtMessenger;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.SetDefaultFont(new Font("Segoe UI", 10F));
        Application.Run(new MainForm());
    }
}

public sealed class Subscriber
{
    public bool Selected { get; set; }
    public string Name { get; set; } = "";
    public string Phone { get; set; } = "";
    public decimal Remaining { get; set; }
    public DateTime? OpenedAt { get; set; }
    public string Status => OpenedAt.HasValue ? "تم فتح واتساب" : "غير مرسل";
}

public sealed class AppState
{
    public List<Subscriber> Subscribers { get; set; } = new();
    public string MessageTemplate { get; set; } =
        "السلام عليكم {name}،\r\nنذكّركم بأن المبلغ المتبقي عليكم لدى تاتلي فون من اشتراك الإنترنت هو {remaining} د.ع.\r\nيرجى التسديد بأقرب وقت ممكن، مع الشكر.";
    public int DelaySeconds { get; set; } = 5;
}

internal static class LocalStore
{
    private static readonly string Folder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Tatli Debt Messenger");
    private static readonly string FilePath = Path.Combine(Folder, "data.json");

    public static AppState Load()
    {
        try
        {
            Directory.CreateDirectory(Folder);
            if (!File.Exists(FilePath)) return new AppState();
            return JsonSerializer.Deserialize<AppState>(File.ReadAllText(FilePath),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new AppState();
        }
        catch
        {
            return new AppState();
        }
    }

    public static void Save(AppState state)
    {
        Directory.CreateDirectory(Folder);
        string temp = FilePath + ".tmp";
        File.WriteAllText(temp, JsonSerializer.Serialize(state,
            new JsonSerializerOptions { WriteIndented = true }));
        File.Move(temp, FilePath, true);
    }
}

public sealed class MainForm : Form
{
    private readonly Color Indigo = Color.FromArgb(43, 45, 104);
    private readonly Color Orange = Color.FromArgb(245, 124, 34);
    private readonly Color Soft = Color.FromArgb(246, 247, 251);
    private readonly AppState state;
    private BindingList<Subscriber> view = new();
    private readonly DataGridView grid = new();
    private readonly TextBox searchBox = new();
    private readonly ComboBox filterBox = new();
    private readonly TextBox templateBox = new();
    private readonly NumericUpDown delayBox = new();
    private readonly Label totalsLabel = new();
    private readonly Label statusLabel = new();
    private readonly ProgressBar progress = new();
    private readonly Button sendButton = new();
    private readonly Button stopButton = new();
    private CancellationTokenSource? sendingCts;
    private readonly System.Windows.Forms.Timer saveTimer = new();

    public MainForm()
    {
        state = LocalStore.Load();
        Text = "Tatli Debt Messenger";
        MinimumSize = new Size(1100, 720);
        Size = new Size(1380, 850);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Soft;
        RightToLeft = RightToLeft.Yes;
        RightToLeftLayout = true;
        Icon = SystemIcons.Information;

        BuildUi();
        ApplyFilter();

        saveTimer.Interval = 650;
        saveTimer.Tick += (_, _) =>
        {
            saveTimer.Stop();
            state.MessageTemplate = templateBox.Text;
            state.DelaySeconds = (int)delayBox.Value;
            SaveNow();
        };
        FormClosing += (_, _) => SaveNow();
    }

    private void BuildUi()
    {
        var header = new Panel { Dock = DockStyle.Top, Height = 86, BackColor = Indigo, Padding = new Padding(25, 12, 25, 10) };
        var title = new Label
        {
            Text = "Tatli Debt Messenger",
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 20F, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(25, 10)
        };
        var subtitle = new Label
        {
            Text = "إدارة ديون المشتركين ورسائل واتساب — تاتلي فون",
            ForeColor = Color.FromArgb(220, 220, 240),
            Font = new Font("Segoe UI", 10F),
            AutoSize = true,
            Location = new Point(29, 51)
        };
        header.Controls.Add(title);
        header.Controls.Add(subtitle);
        Controls.Add(header);

        var toolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 62,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            Padding = new Padding(16, 11, 16, 8),
            BackColor = Color.White
        };
        toolbar.Controls.Add(MakeButton("استيراد Excel", Orange, ImportExcel));
        toolbar.Controls.Add(MakeButton("تنزيل قالب Excel", Indigo, DownloadTemplate));
        toolbar.Controls.Add(MakeButton("إضافة مشترك", Color.FromArgb(22, 145, 101), AddSubscriber));
        toolbar.Controls.Add(MakeButton("حذف المحدد", Color.FromArgb(190, 55, 55), DeleteSelected));
        toolbar.Controls.Add(MakeButton("تحديد الظاهر", Color.FromArgb(88, 96, 125), SelectVisible));
        toolbar.Controls.Add(MakeButton("إلغاء التحديد", Color.FromArgb(120, 125, 140), ClearSelection));
        Controls.Add(toolbar);

        var content = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            SplitterDistance = 880,
            BackColor = Soft,
            Padding = new Padding(14)
        };
        content.Panel1.Padding = new Padding(0, 0, 0, 0);
        content.Panel2.Padding = new Padding(12, 0, 0, 0);
        Controls.Add(content);
        content.BringToFront();

        var left = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(12) };
        content.Panel1.Controls.Add(left);

        var filters = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 54,
            ColumnCount = 2,
            RowCount = 1,
            Padding = new Padding(0, 4, 0, 7)
        };
        filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));
        filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));
        searchBox.Dock = DockStyle.Fill;
        searchBox.PlaceholderText = "بحث بالاسم أو رقم الهاتف...";
        searchBox.BorderStyle = BorderStyle.FixedSingle;
        searchBox.TextChanged += (_, _) => ApplyFilter();
        filterBox.Dock = DockStyle.Fill;
        filterBox.DropDownStyle = ComboBoxStyle.DropDownList;
        filterBox.Items.AddRange(new object[] { "الكل", "عليهم دين", "غير مرسل", "تم فتح واتساب" });
        filterBox.SelectedIndex = 0;
        filterBox.SelectedIndexChanged += (_, _) => ApplyFilter();
        filters.Controls.Add(searchBox, 0, 0);
        filters.Controls.Add(filterBox, 1, 0);
        left.Controls.Add(filters);

        ConfigureGrid();
        left.Controls.Add(grid);
        grid.BringToFront();

        var footer = new Panel { Dock = DockStyle.Bottom, Height = 47, BackColor = Color.FromArgb(250, 250, 252), Padding = new Padding(10) };
        totalsLabel.Dock = DockStyle.Fill;
        totalsLabel.TextAlign = ContentAlignment.MiddleRight;
        totalsLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        totalsLabel.ForeColor = Indigo;
        footer.Controls.Add(totalsLabel);
        left.Controls.Add(footer);

        var right = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(16) };
        content.Panel2.Controls.Add(right);

        var messageTitle = new Label
        {
            Text = "قالب رسالة المطالبة",
            Dock = DockStyle.Top,
            Height = 34,
            Font = new Font("Segoe UI", 12F, FontStyle.Bold),
            ForeColor = Indigo
        };
        right.Controls.Add(messageTitle);

        var hint = new Label
        {
            Text = "المتغيرات: {name} الاسم — {remaining} المبلغ — {phone} الرقم",
            Dock = DockStyle.Top,
            Height = 45,
            ForeColor = Color.DimGray
        };
        right.Controls.Add(hint);
        hint.BringToFront();

        templateBox.Multiline = true;
        templateBox.ScrollBars = ScrollBars.Vertical;
        templateBox.Height = 220;
        templateBox.Dock = DockStyle.Top;
        templateBox.Text = state.MessageTemplate;
        templateBox.Font = new Font("Segoe UI", 10.5F);
        templateBox.TextChanged += (_, _) => ScheduleSave();
        right.Controls.Add(templateBox);
        templateBox.BringToFront();

        var options = new TableLayoutPanel { Dock = DockStyle.Top, Height = 60, ColumnCount = 2, Padding = new Padding(0, 14, 0, 4) };
        options.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62));
        options.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
        options.Controls.Add(new Label { Text = "الفاصل بين الرسائل (ثانية):", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight }, 0, 0);
        delayBox.Minimum = 3;
        delayBox.Maximum = 60;
        delayBox.Value = Math.Clamp(state.DelaySeconds, 3, 60);
        delayBox.Dock = DockStyle.Fill;
        delayBox.ValueChanged += (_, _) => ScheduleSave();
        options.Controls.Add(delayBox, 1, 0);
        right.Controls.Add(options);
        options.BringToFront();

        sendButton.Text = "بدء الإرسال المتسلسل";
        sendButton.Dock = DockStyle.Top;
        sendButton.Height = 48;
        StyleButton(sendButton, Orange);
        sendButton.Click += async (_, _) => await SendSelectedAsync();
        right.Controls.Add(sendButton);
        sendButton.BringToFront();

        stopButton.Text = "إيقاف";
        stopButton.Dock = DockStyle.Top;
        stopButton.Height = 42;
        stopButton.Enabled = false;
        StyleButton(stopButton, Color.FromArgb(180, 58, 58));
        stopButton.Click += (_, _) => sendingCts?.Cancel();
        right.Controls.Add(stopButton);
        stopButton.BringToFront();

        progress.Dock = DockStyle.Top;
        progress.Height = 18;
        progress.Margin = new Padding(0, 12, 0, 0);
        right.Controls.Add(progress);
        progress.BringToFront();

        statusLabel.Text = "جاهز";
        statusLabel.Dock = DockStyle.Top;
        statusLabel.Height = 54;
        statusLabel.TextAlign = ContentAlignment.MiddleCenter;
        statusLabel.ForeColor = Color.DimGray;
        right.Controls.Add(statusLabel);
        statusLabel.BringToFront();

        var note = new Label
        {
            Text = "ملاحظة: يفتح البرنامج محادثة واتساب مع النص جاهزاً. اضغط زر الإرسال داخل واتساب، ثم ينتقل تلقائياً للمشترك التالي حسب الفاصل المحدد.",
            Dock = DockStyle.Bottom,
            Height = 100,
            ForeColor = Color.FromArgb(95, 96, 110),
            BackColor = Color.FromArgb(249, 246, 240),
            Padding = new Padding(12),
            TextAlign = ContentAlignment.MiddleRight
        };
        right.Controls.Add(note);
    }

    private void ConfigureGrid()
    {
        grid.Dock = DockStyle.Fill;
        grid.BackgroundColor = Color.White;
        grid.BorderStyle = BorderStyle.None;
        grid.AutoGenerateColumns = false;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.MultiSelect = false;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.ColumnHeadersHeight = 42;
        grid.RowTemplate.Height = 40;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersDefaultCellStyle.BackColor = Indigo;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 249, 252);

        grid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            DataPropertyName = nameof(Subscriber.Selected),
            HeaderText = "تحديد",
            FillWeight = 43
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(Subscriber.Name),
            HeaderText = "اسم المشترك",
            FillWeight = 155
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(Subscriber.Phone),
            HeaderText = "رقم الهاتف",
            FillWeight = 105
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(Subscriber.Remaining),
            HeaderText = "المبلغ المتبقي",
            FillWeight = 85,
            DefaultCellStyle = new DataGridViewCellStyle { Format = "N0" }
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(Subscriber.Status),
            HeaderText = "الحالة",
            ReadOnly = true,
            FillWeight = 95
        });
        grid.Columns.Add(new DataGridViewButtonColumn
        {
            Name = "WhatsApp",
            HeaderText = "",
            Text = "واتساب",
            UseColumnTextForButtonValue = true,
            FillWeight = 65,
            FlatStyle = FlatStyle.Flat
        });

        grid.CellContentClick += GridCellContentClick;
        grid.CellEndEdit += (_, _) =>
        {
            foreach (var item in state.Subscribers)
                item.Phone = PhoneTools.Normalize(item.Phone);
            SaveNow();
            UpdateTotals();
        };
        grid.CurrentCellDirtyStateChanged += (_, _) =>
        {
            if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };
        grid.CellValueChanged += (_, e) =>
        {
            if (e.RowIndex >= 0) { SaveNow(); UpdateTotals(); }
        };
        grid.DataError += (_, _) => { };
    }

    private Button MakeButton(string text, Color color, EventHandler handler)
    {
        var button = new Button { Text = text, AutoSize = true, Height = 38, Padding = new Padding(13, 0, 13, 0), Margin = new Padding(5, 0, 5, 0) };
        StyleButton(button, color);
        button.Click += handler;
        return button;
    }

    private static void StyleButton(Button button, Color color)
    {
        button.BackColor = color;
        button.ForeColor = Color.White;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.Cursor = Cursors.Hand;
        button.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
    }

    private void ApplyFilter()
    {
        string q = searchBox.Text.Trim();
        string filter = filterBox.SelectedItem?.ToString() ?? "الكل";
        IEnumerable<Subscriber> items = state.Subscribers;

        if (q.Length > 0)
            items = items.Where(x => x.Name.Contains(q, StringComparison.OrdinalIgnoreCase)
                || x.Phone.Contains(q, StringComparison.OrdinalIgnoreCase));

        items = filter switch
        {
            "عليهم دين" => items.Where(x => x.Remaining > 0),
            "غير مرسل" => items.Where(x => !x.OpenedAt.HasValue),
            "تم فتح واتساب" => items.Where(x => x.OpenedAt.HasValue),
            _ => items
        };

        view = new BindingList<Subscriber>(items.ToList());
        grid.DataSource = view;
        UpdateTotals();
    }

    private void UpdateTotals()
    {
        decimal total = view.Sum(x => x.Remaining);
        int selected = state.Subscribers.Count(x => x.Selected);
        totalsLabel.Text = string.Format(CultureInfo.InvariantCulture,
            "الظاهر: {0} مشترك    |    المحدد: {1}    |    مجموع المتبقي: {2:N0} د.ع",
            view.Count, selected, total);
    }

    private void ImportExcel(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "ملفات Excel (*.xlsx)|*.xlsx|ملفات CSV (*.csv)|*.csv",
            Title = "اختر ملف ديون المشتركين"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;

        try
        {
            List<Subscriber> imported = Path.GetExtension(dialog.FileName).Equals(".csv", StringComparison.OrdinalIgnoreCase)
                ? ExcelTools.ReadCsv(dialog.FileName)
                : ExcelTools.ReadXlsx(dialog.FileName);

            int added = 0, updated = 0, skipped = 0;
            foreach (var item in imported)
            {
                item.Phone = PhoneTools.Normalize(item.Phone);
                if (string.IsNullOrWhiteSpace(item.Name) || string.IsNullOrWhiteSpace(item.Phone))
                {
                    skipped++;
                    continue;
                }

                var existing = state.Subscribers.FirstOrDefault(x => PhoneTools.Normalize(x.Phone) == item.Phone);
                if (existing != null)
                {
                    existing.Name = item.Name;
                    existing.Remaining = item.Remaining;
                    updated++;
                }
                else
                {
                    state.Subscribers.Add(item);
                    added++;
                }
            }

            SaveNow();
            ApplyFilter();
            MessageBox.Show(this,
                string.Format("تم الاستيراد بنجاح.\n\nمضاف: {0}\nمحدّث: {1}\nمتروك: {2}", added, updated, skipped),
                "Tatli Debt Messenger", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "تعذر استيراد الملف:\n" + ex.Message +
                "\n\nتأكد أن الأعمدة هي: اسم المشترك، رقم الهاتف، المبلغ المتبقي.",
                "خطأ في الاستيراد", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void DownloadTemplate(object? sender, EventArgs e)
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "ملف Excel (*.xlsx)|*.xlsx",
            FileName = "قالب ديون المشتركين - تاتلي فون.xlsx",
            Title = "حفظ قالب Excel"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            ExcelTools.WriteTemplate(dialog.FileName);
            MessageBox.Show(this, "تم تنزيل قالب Excel بنجاح.", "تم", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "تعذر حفظ القالب:\n" + ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AddSubscriber(object? sender, EventArgs e)
    {
        using var dialog = new SubscriberDialog();
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var item = dialog.Result;
        item.Phone = PhoneTools.Normalize(item.Phone);
        if (string.IsNullOrWhiteSpace(item.Phone))
        {
            MessageBox.Show(this, "رقم الهاتف غير صالح.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        state.Subscribers.Add(item);
        SaveNow();
        ApplyFilter();
    }

    private void DeleteSelected(object? sender, EventArgs e)
    {
        var selected = state.Subscribers.Where(x => x.Selected).ToList();
        if (selected.Count == 0)
        {
            MessageBox.Show(this, "حدد المشتركين المطلوب حذفهم أولاً.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (MessageBox.Show(this, "هل تريد حذف " + selected.Count + " مشترك؟", "تأكيد الحذف",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        foreach (var item in selected) state.Subscribers.Remove(item);
        SaveNow();
        ApplyFilter();
    }

    private void SelectVisible(object? sender, EventArgs e)
    {
        foreach (var item in view) item.Selected = true;
        grid.Refresh();
        SaveNow();
        UpdateTotals();
    }

    private void ClearSelection(object? sender, EventArgs e)
    {
        foreach (var item in state.Subscribers) item.Selected = false;
        grid.Refresh();
        SaveNow();
        UpdateTotals();
    }

    private void GridCellContentClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || grid.Columns[e.ColumnIndex].Name != "WhatsApp") return;
        if (grid.Rows[e.RowIndex].DataBoundItem is not Subscriber item) return;
        OpenWhatsApp(item);
    }

    private void OpenWhatsApp(Subscriber item)
    {
        try
        {
            string phone = PhoneTools.Normalize(item.Phone);
            if (string.IsNullOrWhiteSpace(phone))
                throw new InvalidOperationException("رقم الهاتف غير صالح.");
            string message = BuildMessage(item);
            string url = "https://wa.me/" + phone + "?text=" + Uri.EscapeDataString(message);
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            item.OpenedAt = DateTime.Now;
            SaveNow();
            grid.Refresh();
            UpdateTotals();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "تعذر فتح واتساب:\n" + ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private string BuildMessage(Subscriber item)
    {
        return templateBox.Text
            .Replace("{name}", item.Name, StringComparison.OrdinalIgnoreCase)
            .Replace("{remaining}", item.Remaining.ToString("N0", CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase)
            .Replace("{phone}", PhoneTools.Normalize(item.Phone), StringComparison.OrdinalIgnoreCase);
    }

    private async Task SendSelectedAsync()
    {
        var items = state.Subscribers.Where(x => x.Selected && x.Remaining > 0).ToList();
        if (items.Count == 0)
        {
            MessageBox.Show(this, "حدد مشتركاً واحداً على الأقل وعليه مبلغ متبقٍ.", "تنبيه",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (MessageBox.Show(this,
            "سيتم فتح " + items.Count + " محادثة بالتسلسل.\nاضغط إرسال داخل واتساب لكل رسالة قبل الانتقال للتالية.",
            "بدء الإرسال المتسلسل", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;

        sendingCts = new CancellationTokenSource();
        sendButton.Enabled = false;
        stopButton.Enabled = true;
        progress.Minimum = 0;
        progress.Maximum = items.Count;
        progress.Value = 0;

        try
        {
            for (int i = 0; i < items.Count; i++)
            {
                sendingCts.Token.ThrowIfCancellationRequested();
                statusLabel.Text = string.Format("فتح رسالة {0} من {1}: {2}", i + 1, items.Count, items[i].Name);
                OpenWhatsApp(items[i]);
                progress.Value = i + 1;
                if (i < items.Count - 1)
                    await Task.Delay(TimeSpan.FromSeconds((double)delayBox.Value), sendingCts.Token);
            }
            statusLabel.Text = "اكتمل فتح الرسائل المحددة";
            MessageBox.Show(this, "اكتملت قائمة الإرسال المتسلسل.", "تم", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (OperationCanceledException)
        {
            statusLabel.Text = "تم إيقاف الإرسال";
        }
        finally
        {
            sendingCts.Dispose();
            sendingCts = null;
            sendButton.Enabled = true;
            stopButton.Enabled = false;
            SaveNow();
            ApplyFilter();
        }
    }

    private void ScheduleSave()
    {
        saveTimer.Stop();
        saveTimer.Start();
    }

    private void SaveNow()
    {
        try
        {
            state.MessageTemplate = templateBox.Text;
            state.DelaySeconds = (int)delayBox.Value;
            LocalStore.Save(state);
        }
        catch
        {
            statusLabel.Text = "تعذر حفظ البيانات محلياً";
        }
    }
}

public sealed class SubscriberDialog : Form
{
    private readonly TextBox name = new();
    private readonly TextBox phone = new();
    private readonly NumericUpDown amount = new();
    public Subscriber Result { get; private set; } = new();

    public SubscriberDialog()
    {
        Text = "إضافة مشترك";
        Size = new Size(470, 310);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        RightToLeft = RightToLeft.Yes;
        RightToLeftLayout = true;

        var table = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), RowCount = 4, ColumnCount = 2 };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
        table.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        table.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        table.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        table.Controls.Add(new Label { Text = "اسم المشترك", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight }, 0, 0);
        name.Dock = DockStyle.Fill;
        table.Controls.Add(name, 1, 0);
        table.Controls.Add(new Label { Text = "رقم الهاتف", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight }, 0, 1);
        phone.Dock = DockStyle.Fill;
        table.Controls.Add(phone, 1, 1);
        table.Controls.Add(new Label { Text = "المبلغ المتبقي", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight }, 0, 2);
        amount.Dock = DockStyle.Fill;
        amount.Maximum = 1000000000;
        amount.ThousandsSeparator = true;
        table.Controls.Add(amount, 1, 2);

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var ok = new Button { Text = "حفظ", DialogResult = DialogResult.None, Width = 110, Height = 38, BackColor = Color.FromArgb(245, 124, 34), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        var cancel = new Button { Text = "إلغاء", DialogResult = DialogResult.Cancel, Width = 90, Height = 38 };
        ok.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(name.Text) || string.IsNullOrWhiteSpace(phone.Text))
            {
                MessageBox.Show(this, "أدخل اسم المشترك ورقم الهاتف.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Result = new Subscriber { Name = name.Text.Trim(), Phone = phone.Text.Trim(), Remaining = amount.Value };
            DialogResult = DialogResult.OK;
            Close();
        };
        buttons.Controls.Add(ok);
        buttons.Controls.Add(cancel);
        table.Controls.Add(buttons, 0, 3);
        table.SetColumnSpan(buttons, 2);
        Controls.Add(table);
        AcceptButton = ok;
        CancelButton = cancel;
    }
}

internal static class PhoneTools
{
    public static string Normalize(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        string phone = ConvertArabicDigits(value);
        phone = Regex.Replace(phone, @"[^\d]", "");

        if (phone.StartsWith("00964")) phone = phone[2..];
        if (phone.StartsWith("9640")) phone = "964" + phone[4..];
        else if (phone.StartsWith("07")) phone = "964" + phone[1..];
        else if (phone.StartsWith("7") && phone.Length == 10) phone = "964" + phone;

        return phone;
    }

    private static string ConvertArabicDigits(string input)
    {
        const string ar = "٠١٢٣٤٥٦٧٨٩";
        const string fa = "۰۱۲۳۴۵۶۷۸۹";
        var sb = new StringBuilder(input.Length);
        foreach (char c in input)
        {
            int i = ar.IndexOf(c);
            if (i >= 0) { sb.Append((char)('0' + i)); continue; }
            i = fa.IndexOf(c);
            sb.Append(i >= 0 ? (char)('0' + i) : c);
        }
        return sb.ToString();
    }
}

internal static class ExcelTools
{
    private static readonly string[] NameHeaders = { "اسم المشترك", "الاسم", "اسم", "name", "subscriber", "subscriber name" };
    private static readonly string[] PhoneHeaders = { "رقم الهاتف", "الهاتف", "رقم الموبايل", "الموبايل", "phone", "mobile", "number" };
    private static readonly string[] AmountHeaders = { "المبلغ المتبقي", "المتبقي", "الدين", "المبلغ", "remaining", "balance", "debt", "amount" };

    public static List<Subscriber> ReadXlsx(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var sharedStrings = ReadSharedStrings(archive);
        var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")
            ?? archive.Entries.FirstOrDefault(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase));
        if (sheetEntry == null) throw new InvalidDataException("لا توجد ورقة عمل داخل الملف.");

        XDocument doc;
        using (var stream = sheetEntry.Open()) doc = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var rows = new List<List<string>>();

        foreach (var row in doc.Descendants(ns + "row"))
        {
            var values = new SortedDictionary<int, string>();
            foreach (var cell in row.Elements(ns + "c"))
            {
                string reference = (string?)cell.Attribute("r") ?? "";
                int col = ColumnIndex(reference);
                string type = (string?)cell.Attribute("t") ?? "";
                string raw = cell.Element(ns + "v")?.Value ?? "";
                string value;
                if (type == "s" && int.TryParse(raw, out int si) && si >= 0 && si < sharedStrings.Count)
                    value = sharedStrings[si];
                else if (type == "inlineStr")
                    value = string.Concat(cell.Descendants(ns + "t").Select(x => x.Value));
                else
                    value = NormalizeExcelNumber(raw);
                values[col] = value.Trim();
            }

            if (values.Count > 0)
            {
                int max = values.Keys.Max();
                var line = Enumerable.Repeat("", max + 1).ToList();
                foreach (var pair in values) line[pair.Key] = pair.Value;
                rows.Add(line);
            }
        }
        return RowsToSubscribers(rows);
    }

    public static List<Subscriber> ReadCsv(string path)
    {
        string[] lines = File.ReadAllLines(path, Encoding.UTF8);
        if (lines.Length == 0) return new List<Subscriber>();
        char delimiter = lines[0].Count(c => c == ';') > lines[0].Count(c => c == ',') ? ';' : ',';
        var rows = lines.Select(line => ParseCsvLine(line, delimiter)).ToList();
        return RowsToSubscribers(rows);
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return new List<string>();
        XDocument doc;
        using (var stream = entry.Open()) doc = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        return doc.Descendants(ns + "si")
            .Select(si => string.Concat(si.Descendants(ns + "t").Select(t => t.Value)))
            .ToList();
    }

    private static List<Subscriber> RowsToSubscribers(List<List<string>> rows)
    {
        if (rows.Count == 0) return new List<Subscriber>();
        int headerRow = rows.FindIndex(r => r.Any(v => HeaderMatches(v, NameHeaders))
            && r.Any(v => HeaderMatches(v, PhoneHeaders)));
        if (headerRow < 0) throw new InvalidDataException("لم يتم العثور على عناوين الأعمدة المطلوبة.");

        var headers = rows[headerRow];
        int nameCol = FindHeader(headers, NameHeaders);
        int phoneCol = FindHeader(headers, PhoneHeaders);
        int amountCol = FindHeader(headers, AmountHeaders);
        if (nameCol < 0 || phoneCol < 0 || amountCol < 0)
            throw new InvalidDataException("يجب أن يحتوي الملف على: اسم المشترك، رقم الهاتف، المبلغ المتبقي.");

        var result = new List<Subscriber>();
        foreach (var row in rows.Skip(headerRow + 1))
        {
            string name = Get(row, nameCol).Trim();
            string phone = Get(row, phoneCol).Trim();
            string amountText = Get(row, amountCol).Trim();
            if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(phone)) continue;
            decimal amount = ParseAmount(amountText);
            result.Add(new Subscriber { Name = name, Phone = phone, Remaining = amount });
        }
        return result;
    }

    private static decimal ParseAmount(string value)
    {
        value = value.Replace("د.ع", "").Replace("دينار", "").Trim();
        if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal number)) return number;
        if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out number)) return number;
        value = value.Replace(",", "");
        return decimal.TryParse(value, out number) ? number : 0;
    }

    private static int FindHeader(List<string> row, string[] options)
        => row.FindIndex(value => HeaderMatches(value, options));

    private static bool HeaderMatches(string value, string[] options)
    {
        string normalized = Regex.Replace(value.Trim().ToLowerInvariant(), @"\s+", " ");
        return options.Any(x => normalized == x);
    }

    private static string Get(List<string> row, int index) => index >= 0 && index < row.Count ? row[index] : "";

    private static int ColumnIndex(string reference)
    {
        int result = 0;
        foreach (char c in reference)
        {
            if (!char.IsLetter(c)) break;
            result = result * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }
        return Math.Max(0, result - 1);
    }

    private static string NormalizeExcelNumber(string raw)
    {
        if (raw.IndexOfAny(new[] { 'E', 'e' }) >= 0
            && decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal d))
            return d.ToString("0", CultureInfo.InvariantCulture);
        return raw;
    }

    private static List<string> ParseCsvLine(string line, char delimiter)
    {
        var result = new List<string>();
        var current = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"')
            {
                if (quoted && i + 1 < line.Length && line[i + 1] == '"') { current.Append('"'); i++; }
                else quoted = !quoted;
            }
            else if (c == delimiter && !quoted)
            {
                result.Add(current.ToString());
                current.Clear();
            }
            else current.Append(c);
        }
        result.Add(current.ToString());
        return result;
    }

    public static void WriteTemplate(string path)
    {
        if (File.Exists(path)) File.Delete(path);
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Default Extension=""xml"" ContentType=""application/xml""/>
<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>");
        Add(archive, "_rels/.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");
        Add(archive, "xl/workbook.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets><sheet name=""ديون المشتركين"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>");
        Add(archive, "xl/_rels/workbook.xml.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
<Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>");
        Add(archive, "xl/styles.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<fonts count=""2""><font><sz val=""11""/><name val=""Calibri""/></font><font><b/><color rgb=""FFFFFFFF""/><sz val=""12""/><name val=""Arial""/></font></fonts>
<fills count=""3""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FF2B2D68""/><bgColor indexed=""64""/></patternFill></fill></fills>
<borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>
<cellXfs count=""2""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/><xf numFmtId=""0"" fontId=""1"" fillId=""2"" borderId=""0"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyAlignment=""1""><alignment horizontal=""center"" readingOrder=""2""/></xf></cellXfs>
<cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>
</styleSheet>");
        Add(archive, "xl/worksheets/sheet1.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<sheetViews><sheetView rightToLeft=""1"" workbookViewId=""0""/></sheetViews>
<cols><col min=""1"" max=""1"" width=""28"" customWidth=""1""/><col min=""2"" max=""2"" width=""22"" customWidth=""1""/><col min=""3"" max=""3"" width=""20"" customWidth=""1""/></cols>
<sheetData>
<row r=""1"" ht=""25"" customHeight=""1"">
<c r=""A1"" t=""inlineStr"" s=""1""><is><t>اسم المشترك</t></is></c>
<c r=""B1"" t=""inlineStr"" s=""1""><is><t>رقم الهاتف</t></is></c>
<c r=""C1"" t=""inlineStr"" s=""1""><is><t>المبلغ المتبقي</t></is></c>
</row>
<row r=""2"">
<c r=""A2"" t=""inlineStr""><is><t>مثال: أحمد علي</t></is></c>
<c r=""B2"" t=""inlineStr""><is><t>07814116000</t></is></c>
<c r=""C2""><v>25000</v></c>
</row>
</sheetData>
<autoFilter ref=""A1:C10000""/>
</worksheet>");
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
