// «Менеджер плагинов КОМПАС» — версия, настройки (settings.ini), обновление с GitHub.
// Сборка: _setup\manager\build.cmd (csc из .NET Framework 4.x, есть в любой Windows 10/11).
// Окно ничего не знает о плагинах заранее: список берётся из папок с plugin.ini,
// поля настроек — из _setup\settings_schema.json, обновление делает _setup\update.py.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("Менеджер плагинов КОМПАС")]
[assembly: System.Reflection.AssemblyProduct("Плагины КОМПАС")]

namespace KompasPlugins
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string tab = null, plugin = "", errorFile = null, title = null, section = "frame";
            bool report = false, askSides = false;
            foreach (var a in args)
            {
                if (a.StartsWith("--tab=")) tab = a.Substring(6);
                else if (a == "--report") report = true;
                else if (a == "--ask-sides") askSides = true;
                else if (a.StartsWith("--section=")) section = a.Substring(10);
                else if (a.StartsWith("--plugin=")) plugin = a.Substring(9);
                else if (a.StartsWith("--error=")) errorFile = a.Substring(8);
                else if (a.StartsWith("--title=")) title = a.Substring(8);
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var root = Path.GetDirectoryName(Application.ExecutablePath);
            if (askSides)
            {
                // окно «Какие виды строить?» → код выхода: 1 — слева, 2 — справа, 3 — оба, 0 — отмена
                var f = new SidesForm(root, section, title);
                Application.Run(f);
                Environment.ExitCode = f.Result;
            }
            else if (report)
                Application.Run(new ReportForm(root, plugin, errorFile, title));
            else
                Application.Run(new MainForm(tab));
        }
    }

    // ---------- настройки: схема + settings.ini ----------
    class SettingItem
    {
        public string Section, Key, Type, Label, Help;
        public object Default;
        public List<KeyValuePair<string, string>> Choices = new List<KeyValuePair<string, string>>();
        public double Min = 0, Max = 1000;
    }

    class SettingSection
    {
        public string Id, Title;
        public List<SettingItem> Items = new List<SettingItem>();
    }

    class SettingsStore
    {
        public readonly string Root, File, SchemaFile;
        public List<SettingSection> Sections = new List<SettingSection>();
        public Dictionary<string, string> Values = new Dictionary<string, string>(); // "section.key" → текст

        public SettingsStore(string root)
        {
            Root = root;
            File = Path.Combine(root, "settings.ini");
            SchemaFile = Path.Combine(root, "_setup", "settings_schema.json");
            LoadSchema();
            LoadValues();
        }

        void LoadSchema()
        {
            var js = new JavaScriptSerializer();
            var data = (Dictionary<string, object>)js.DeserializeObject(System.IO.File.ReadAllText(SchemaFile, Encoding.UTF8));
            foreach (Dictionary<string, object> s in (object[])data["sections"])
            {
                var sec = new SettingSection { Id = (string)s["id"], Title = (string)s["title"] };
                foreach (Dictionary<string, object> it in (object[])s["items"])
                {
                    var item = new SettingItem
                    {
                        Section = sec.Id, Key = (string)it["key"], Type = (string)it["type"],
                        Label = (string)it["label"], Default = it["default"],
                        Help = it.ContainsKey("help") ? (string)it["help"] : null
                    };
                    if (it.ContainsKey("min")) item.Min = Convert.ToDouble(it["min"]);
                    if (it.ContainsKey("max")) item.Max = Convert.ToDouble(it["max"]);
                    if (it.ContainsKey("choices"))
                        foreach (object[] c in (object[])it["choices"])
                            item.Choices.Add(new KeyValuePair<string, string>((string)c[0], (string)c[1]));
                    sec.Items.Add(item);
                }
                Sections.Add(sec);
            }
        }

        public void LoadValues()
        {
            Values.Clear();
            if (!System.IO.File.Exists(File)) return;
            string section = "";
            foreach (var raw in System.IO.File.ReadAllLines(File, Encoding.UTF8))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith(";") || line.StartsWith("#")) continue;
                if (line.StartsWith("[") && line.EndsWith("]")) { section = line.Substring(1, line.Length - 2).Trim(); continue; }
                int eq = line.IndexOf('=');
                if (eq < 0) continue;
                var val = line.Substring(eq + 1);
                int sc = val.IndexOf(" ;");
                if (sc >= 0) val = val.Substring(0, sc);
                Values[section + "." + line.Substring(0, eq).Trim()] = val.Trim();
            }
        }

        public static string DefaultText(SettingItem it)
        {
            if (it.Default is bool) return (bool)it.Default ? "true" : "false";
            if (it.Default is decimal || it.Default is double || it.Default is int)
                return Convert.ToDouble(it.Default).ToString(System.Globalization.CultureInfo.InvariantCulture);
            return it.Default == null ? "" : it.Default.ToString();
        }

        public string Get(SettingItem it)
        {
            string v;
            return Values.TryGetValue(it.Section + "." + it.Key, out v) ? v : DefaultText(it);
        }

        public void Set(SettingItem it, string value) { Values[it.Section + "." + it.Key] = value; }

        public void Save()
        {
            var sb = new StringBuilder();
            sb.AppendLine("; Настройки плагинов КОМПАС. Правятся в «Менеджер плагинов КОМПАС.exe» (можно и Блокнотом).");
            sb.AppendLine("; Обновление с GitHub этот файл не трогает. Удалите строку — вернётся значение по умолчанию.");
            foreach (var s in Sections)
            {
                sb.AppendLine();
                sb.AppendLine("[" + s.Id + "]");
                sb.AppendLine("; " + s.Title);
                foreach (var it in s.Items)
                {
                    var comment = it.Label;
                    if (it.Choices.Count > 0)
                        comment += " (" + string.Join(", ", it.Choices.Select(c => c.Key + " — " + c.Value)) + ")";
                    sb.AppendLine("; " + comment);
                    sb.AppendLine(it.Key + " = " + Get(it));
                }
            }
            System.IO.File.WriteAllText(File, sb.ToString(), new UTF8Encoding(false));
        }
    }

    // ---------- главное окно ----------
    class MainForm : Form
    {
        readonly string Root;
        readonly SettingsStore Store;
        TabControl tabs;
        TabPage tabPlugins, tabSettings, tabUpdate, tabAbout;
        ListBox sectionList;
        Panel settingsPanel;
        Label lblInstalled, lblLatest, lblStatus;
        TextBox txtNotes, txtLog;
        Button btnCheck, btnUpdate;
        readonly ToolTip tips = new ToolTip { AutoPopDelay = 15000 };
        readonly Dictionary<SettingItem, Control> editors = new Dictionary<SettingItem, Control>();

        public MainForm(string startTab)
        {
            Root = Path.GetDirectoryName(Application.ExecutablePath);
            Store = new SettingsStore(Root);
            Text = "Менеджер плагинов КОМПАС — версия " + ReadVersion();
            Font = new Font("Segoe UI", 9.5f);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(820, 560);
            MinimumSize = new Size(700, 480);
            StartPosition = FormStartPosition.CenterScreen;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(14, 5) };
            tabPlugins = new TabPage("Плагины");
            tabSettings = new TabPage("Настройки");
            tabUpdate = new TabPage("Версия и обновление");
            tabAbout = new TabPage("Что нового");
            tabs.TabPages.AddRange(new[] { tabPlugins, tabSettings, tabUpdate, tabAbout });
            Controls.Add(tabs);

            BuildPluginsTab();
            BuildSettingsTab();
            BuildUpdateTab();
            BuildAboutTab();

            if (startTab == "update") tabs.SelectedTab = tabUpdate;
            else if (!string.IsNullOrEmpty(startTab))
            {
                tabs.SelectedTab = tabSettings;
                int i = Store.Sections.FindIndex(s => s.Id == startTab);
                if (i >= 0) sectionList.SelectedIndex = sectionList.Items.IndexOf(Store.Sections[i].Title);
            }
            Shown += (s, e) => CheckUpdates(false);
        }

        string ReadVersion()
        {
            var f = Path.Combine(Root, "VERSION");
            return File.Exists(f) ? File.ReadAllText(f).Trim() : "?";
        }

        // ---------- вкладка «Плагины» ----------
        void BuildPluginsTab()
        {
            var head = new Label
            {
                Dock = DockStyle.Top, Height = 48, Padding = new Padding(8, 8, 8, 0),
                Text = "Версия плагинов: " + ReadVersion() + "\r\nПапка: " + Root
            };
            var list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true };
            list.Columns.Add("Плагин (меню «Приложения» в КОМПАС)", 230);
            list.Columns.Add("Что делает", 540);
            foreach (var dir in Directory.GetDirectories(Root).OrderBy(d => d))
            {
                var ini = Path.Combine(dir, "plugin.ini");
                if (!File.Exists(ini)) continue;
                var kv = ReadIni(ini, "plugin");
                string name, desc;
                kv.TryGetValue("name", out name);
                kv.TryGetValue("description", out desc);
                list.Items.Add(new ListViewItem(new[] { name ?? Path.GetFileName(dir), desc ?? "" }));
            }
            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 46, Padding = new Padding(6) };
            var bOpen = new Button { Text = "Открыть папку плагинов", AutoSize = true };
            bOpen.Click += (s, e) => Process.Start("explorer.exe", "\"" + Root + "\"");
            var bReg = new Button { Text = "Подключить в КОМПАС заново", AutoSize = true };
            tips.SetToolTip(bReg, "Если пункт плагина пропал из меню «Приложения». КОМПАС должен быть закрыт.");
            bReg.Click += (s, e) =>
            {
                var inst = Path.Combine(Root, "_setup", "install.py");
                RunPython("\"" + inst + "\" --root \"" + Root + "\"", false, null);
            };
            var bBug = new Button { Text = "Сообщить о проблеме…", AutoSize = true };
            tips.SetToolTip(bBug, "Отправить разработчику журнал работы плагина и комментарий (Telegram)");
            bBug.Click += (s, e) => new ReportForm(Root, "", null, null).ShowDialog(this);
            bar.Controls.AddRange(new Control[] { bOpen, bReg, bBug });
            tabPlugins.Controls.Add(list);
            tabPlugins.Controls.Add(bar);
            tabPlugins.Controls.Add(head);
        }

        // ---------- вкладка «Настройки» ----------
        void BuildSettingsTab()
        {
            sectionList = new ListBox { Dock = DockStyle.Left, Width = 190, IntegralHeight = false, ItemHeight = 22 };
            foreach (var s in Store.Sections)
                if (s.Id != "update") sectionList.Items.Add(s.Title);
            settingsPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(12, 6, 12, 6) };
            sectionList.SelectedIndexChanged += (s, e) => ShowSection();

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 46, Padding = new Padding(6), FlowDirection = FlowDirection.RightToLeft };
            var bSave = new Button { Text = "Сохранить", AutoSize = true };
            bSave.Click += (s, e) => { SaveEditors(); Store.Save(); lblStatusSet("Сохранено: " + Store.File); };
            var bReset = new Button { Text = "По умолчанию (этот раздел)", AutoSize = true };
            bReset.Click += (s, e) =>
            {
                var sec = CurrentSection();
                if (sec == null) return;
                foreach (var it in sec.Items) Store.Values.Remove(it.Section + "." + it.Key);
                ShowSection();
            };
            var bFile = new Button { Text = "Открыть settings.ini", AutoSize = true };
            bFile.Click += (s, e) => { SaveEditors(); Store.Save(); Py.OpenFile(this, Store.File); };
            lblStatus = new Label { AutoSize = true, ForeColor = Color.DimGray, Padding = new Padding(0, 8, 12, 0) };
            bar.Controls.AddRange(new Control[] { bSave, bReset, bFile, lblStatus });

            tabSettings.Controls.Add(settingsPanel);
            tabSettings.Controls.Add(sectionList);
            tabSettings.Controls.Add(bar);
            if (sectionList.Items.Count > 0) sectionList.SelectedIndex = 0;
        }

        void lblStatusSet(string t) { lblStatus.Text = t; }

        SettingSection CurrentSection()
        {
            if (sectionList.SelectedItem == null) return null;
            return Store.Sections.First(s => s.Title == (string)sectionList.SelectedItem);
        }

        void ShowSection()
        {
            SaveEditors();
            settingsPanel.Controls.Clear();
            foreach (var k in editors.Keys.Where(k => k.Section != "update").ToList()) editors.Remove(k);
            var sec = CurrentSection();
            if (sec == null) return;
            var grid = BuildEditors(sec, null);
            settingsPanel.Controls.Add(grid);
            lblStatusSet("");
        }

        TableLayoutPanel BuildEditors(SettingSection sec, Action onChange)
        {
            var grid = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 270));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            foreach (var it in sec.Items)
            {
                Control ed;
                var val = Store.Get(it);
                if (it.Type == "bool")
                {
                    var cb = new CheckBox { Text = it.Label, AutoSize = true, Checked = val.ToLower() == "true" || val == "1" || val.ToLower() == "да" };
                    if (onChange != null) cb.CheckedChanged += (s, e) => onChange();
                    ed = cb;
                    grid.Controls.Add(cb);
                    grid.SetColumnSpan(cb, 2);
                }
                else
                {
                    var lbl = new Label { Text = it.Label, AutoSize = true, Padding = new Padding(0, 6, 0, 0) };
                    if (it.Type == "choice")
                    {
                        var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 330 };
                        foreach (var c in it.Choices) combo.Items.Add(c.Value);
                        int idx = it.Choices.FindIndex(c => c.Key == val);
                        combo.SelectedIndex = idx >= 0 ? idx : Math.Max(0, it.Choices.FindIndex(c => c.Key == SettingsStore.DefaultText(it)));
                        if (onChange != null) combo.SelectedIndexChanged += (s, e) => onChange();
                        ed = combo;
                    }
                    else if (it.Type == "float" || it.Type == "int")
                    {
                        var num = new NumericUpDown
                        {
                            Width = 110, DecimalPlaces = it.Type == "float" ? 2 : 0, Increment = it.Type == "float" ? 0.1m : 1m,
                            Minimum = (decimal)it.Min, Maximum = (decimal)it.Max
                        };
                        double d;
                        if (!double.TryParse(val.Replace(',', '.'), System.Globalization.NumberStyles.Float,
                                             System.Globalization.CultureInfo.InvariantCulture, out d))
                            d = Convert.ToDouble(it.Default);
                        num.Value = Math.Min(num.Maximum, Math.Max(num.Minimum, (decimal)d));
                        if (onChange != null) num.ValueChanged += (s, e) => onChange();
                        ed = num;
                    }
                    else
                    {
                        var tb = new TextBox { Width = 330, Text = val };
                        if (onChange != null) tb.TextChanged += (s, e) => onChange();
                        ed = tb;
                    }
                    grid.Controls.Add(lbl);
                    grid.Controls.Add(ed);
                    if (it.Help != null) tips.SetToolTip(lbl, it.Help);
                }
                if (it.Help != null)
                {
                    tips.SetToolTip(ed, it.Help);
                    var help = new Label { Text = it.Help, AutoSize = true, ForeColor = Color.Gray, Padding = new Padding(it.Type == "bool" ? 18 : 0, 0, 0, 6) };
                    if (it.Type == "bool") { grid.Controls.Add(help); grid.SetColumnSpan(help, 2); }
                    else { grid.Controls.Add(new Label { AutoSize = true }); grid.Controls.Add(help); }
                }
                editors[it] = ed;
            }
            return grid;
        }

        void SaveEditors()
        {
            foreach (var p in editors)
            {
                var it = p.Key; var c = p.Value;
                if (c is CheckBox) Store.Set(it, ((CheckBox)c).Checked ? "true" : "false");
                else if (c is ComboBox) { int i = ((ComboBox)c).SelectedIndex; if (i >= 0) Store.Set(it, it.Choices[i].Key); }
                else if (c is NumericUpDown) Store.Set(it, ((NumericUpDown)c).Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                else Store.Set(it, c.Text);
            }
        }

        // ---------- вкладка «Версия и обновление» ----------
        void BuildUpdateTab()
        {
            var top = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(8) };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            lblInstalled = new Label { AutoSize = true, Font = new Font(Font, FontStyle.Bold), Text = ReadVersion() };
            lblLatest = new Label { AutoSize = true, Text = "проверяется…" };
            top.Controls.Add(new Label { Text = "Установлена версия:", AutoSize = true });
            top.Controls.Add(lblInstalled);
            top.Controls.Add(new Label { Text = "На GitHub:", AutoSize = true });
            top.Controls.Add(lblLatest);

            var upd = Store.Sections.FirstOrDefault(s => s.Id == "update");
            Control updEditors = upd != null ? BuildEditors(upd, () => { SaveEditors(); Store.Save(); }) : new Panel();
            updEditors.Dock = DockStyle.Top;

            var bar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 44, Padding = new Padding(6) };
            btnCheck = new Button { Text = "Проверить", AutoSize = true };
            btnCheck.Click += (s, e) => CheckUpdates(true);
            btnUpdate = new Button { Text = "Обновить сейчас", AutoSize = true, Enabled = false };
            btnUpdate.Click += (s, e) => DoUpdate();
            bar.Controls.AddRange(new Control[] { btnCheck, btnUpdate });

            txtNotes = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = SystemColors.Window };
            txtLog = new TextBox { Dock = DockStyle.Bottom, Height = 110, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Consolas", 9f) };

            tabUpdate.Controls.Add(txtNotes);
            tabUpdate.Controls.Add(new Label { Dock = DockStyle.Top, Text = "Что в новой версии:", Height = 22, Padding = new Padding(8, 4, 0, 0) });
            tabUpdate.Controls.Add(bar);
            tabUpdate.Controls.Add(updEditors);
            tabUpdate.Controls.Add(top);
            tabUpdate.Controls.Add(txtLog);
        }


        void CheckUpdates(bool userClick)
        {
            btnCheck.Enabled = false;
            lblLatest.Text = "проверяется…";
            var upd = Path.Combine(Root, "_setup", "update.py");
            RunPython("\"" + upd + "\" --root \"" + Root + "\" --check", true, output =>
            {
                btnCheck.Enabled = true;
                // версия могла смениться, пока менеджер открыт (обновление при входе в Windows)
                lblInstalled.Text = ReadVersion();
                Text = "Менеджер плагинов КОМПАС — версия " + ReadVersion();
                try
                {
                    var json = output.Split('\n').Last(l => l.Trim().StartsWith("{"));
                    var d = (Dictionary<string, object>)new JavaScriptSerializer().DeserializeObject(json);
                    if (d.ContainsKey("error"))
                    {
                        lblLatest.Text = "нет связи с GitHub";
                        lblLatest.ForeColor = Color.DarkRed;
                        txtNotes.Text = ("Не удалось узнать новую версию. Проверьте интернет и нажмите «Проверить».\n" +
                                         "Если не помогает — вкладка «Плагины» → «Сообщить о проблеме…» " +
                                         "(уйдёт журнал «Обновление» с подробностями).\n\nОшибка:\n" + (string)d["error"])
                                        .Replace("\r\n", "\n").Replace("\n", "\r\n");
                        return;
                    }
                    bool avail = (bool)d["update_available"];
                    bool pending = (bool)d["register_pending"];
                    lblLatest.Text = d["latest_version"] + (avail ? "  — есть обновление" : "  — у вас последняя");
                    lblLatest.ForeColor = avail ? Color.DarkGreen : SystemColors.ControlText;
                    string title = (string)d["title"], notes = (string)d["notes"];
                    string text = notes.StartsWith(title) || notes.Length == 0 ? (notes.Length > 0 ? notes : title) : title + "\n\n" + notes;
                    txtNotes.Text = text.Replace("\r\n", "\n").Replace("\n", "\r\n");
                    btnUpdate.Enabled = avail || pending;
                    btnUpdate.Text = avail ? "Обновить до " + d["latest_version"] : (pending ? "Подключить новые плагины" : "Обновить сейчас");
                }
                catch (Exception ex)
                {
                    lblLatest.Text = "не удалось проверить";
                    if (userClick) txtLog.Text = output + "\r\n" + ex.Message;
                }
            });
        }

        void DoUpdate()
        {
            btnUpdate.Enabled = btnCheck.Enabled = false;
            txtLog.Text = "Обновление…\r\n";
            var upd = Path.Combine(Root, "_setup", "update.py");
            RunPython("\"" + upd + "\" --root \"" + Root + "\" --quiet", true, output =>
            {
                txtLog.Text = output.Replace("\r\n", "\n").Replace("\n", "\r\n");
                lblInstalled.Text = ReadVersion();
                Text = "Менеджер плагинов КОМПАС — версия " + ReadVersion();
                MessageBox.Show(this, output.Trim().Length > 0 ? output.Trim() : "Готово.", "Обновление плагинов",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                CheckUpdates(false);
            });
        }

        // ---------- вкладка «Что нового» ----------
        void BuildAboutTab()
        {
            var f = Path.Combine(Root, "CHANGELOG.md");
            var tb = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = SystemColors.Window,
                Text = File.Exists(f) ? File.ReadAllText(f, Encoding.UTF8).Replace("\r\n", "\n").Replace("\n", "\r\n") : "Нет файла CHANGELOG.md"
            };
            tabAbout.Controls.Add(tb);
        }

        // ---------- общее ----------
        static Dictionary<string, string> ReadIni(string path, string section)
        {
            var kv = new Dictionary<string, string>();
            if (!File.Exists(path)) return kv;
            string cur = "";
            foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
            {
                var line = raw.Trim();
                if (line.StartsWith(";") || line.Length == 0) continue;
                if (line.StartsWith("[")) { cur = line.Trim('[', ']').Trim(); continue; }
                int eq = line.IndexOf('=');
                if (eq > 0 && cur == section) kv[line.Substring(0, eq).Trim()] = line.Substring(eq + 1).Trim();
            }
            return kv;
        }

        void RunPython(string args, bool capture, Action<string> done) { Py.Run(this, Root, args, capture, done); }
    }

    // ---------- запуск Python (общий для окон) ----------
    static class Py
    {
        // Открыть текстовый файл: программой по умолчанию, а если для расширения её нет — Блокнотом.
        public static void OpenFile(Form owner, string path)
        {
            try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); return; }
            catch { }
            try { Process.Start("notepad.exe", "\"" + path + "\""); }
            catch (Exception ex) { MessageBox.Show(owner, "Не удалось открыть файл:\n" + path + "\n\n" + ex.Message); }
        }

        // Показать текстовый файл (журнал) в своём окне — без Блокнота: новый Блокнот Windows 11
        // на путях с кириллицей и пробелами отвечал «Системе не удается найти указанный путь».
        public static void ShowText(Form owner, string path, Action send = null)
        {
            string text;
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sr = new StreamReader(fs, Encoding.UTF8))
                    text = sr.ReadToEnd();
            }
            catch (Exception ex) { MessageBox.Show(owner, "Не удалось прочитать журнал:\n" + path + "\n\n" + ex.Message); return; }
            var f = new Form
            {
                Text = "Журнал — " + Path.GetFileName(path), Font = owner.Font, StartPosition = FormStartPosition.CenterParent,
                ClientSize = new Size(900, 560), ShowIcon = false, MinimizeBox = false
            };
            var tb = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false,
                Font = new Font("Consolas", 9f), BackColor = SystemColors.Window,
                Text = text.Replace("\r\n", "\n").Replace("\n", "\r\n")
            };
            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, Padding = new Padding(6), FlowDirection = FlowDirection.RightToLeft };
            var bClose = new Button { Text = "Закрыть", AutoSize = true, DialogResult = DialogResult.Cancel };
            var bFolder = new Button { Text = "Показать файл в папке", AutoSize = true };
            bFolder.Click += (s, e) => Process.Start("explorer.exe", "/select,\"" + path + "\"");
            var bCopy = new Button { Text = "Копировать всё", AutoSize = true };
            bCopy.Click += (s, e) => { if (tb.Text.Length > 0) Clipboard.SetText(tb.Text); };
            bar.Controls.AddRange(new Control[] { bClose, bFolder, bCopy });
            if (send != null)
            {
                // отправка — из окна отчёта (там комментарий и статус); журнал закрываем
                var bSend = new Button { Text = "Отправить в Telegram", AutoSize = true, Font = new Font(owner.Font, FontStyle.Bold) };
                bSend.Click += (s, e) => { f.Close(); send(); };
                bar.Controls.Add(bSend);
                f.AcceptButton = bSend;
            }
            f.Controls.Add(tb);
            f.Controls.Add(bar);
            f.CancelButton = bClose;
            f.Shown += (s, e) => { tb.SelectionStart = tb.TextLength; tb.ScrollToCaret(); };
            f.ShowDialog(owner);
        }

        public static string Exe(string root, bool console)
        {
            var name = console ? "python.exe" : "pythonw.exe";
            foreach (var dir in new[] { Path.Combine(root, "Python"), @"C:\KompasPlugins\Python",
                                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KompasPlugins", "Python") })
                if (File.Exists(Path.Combine(dir, name))) return Path.Combine(dir, name);
            return name;
        }

        // capture=true — без окна, вывод отдаётся в done (в потоке окна)
        public static void Run(Form owner, string Root, string args, bool capture, Action<string> done)
        {
            var psi = new ProcessStartInfo(Exe(Root, capture), args)
            {
                UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Root,
                RedirectStandardOutput = capture, RedirectStandardError = capture
            };
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            if (capture) { psi.StandardOutputEncoding = Encoding.UTF8; psi.StandardErrorEncoding = Encoding.UTF8; }
            new Thread(() =>
            {
                string output;
                try
                {
                    using (var p = Process.Start(psi))
                    {
                        if (capture)
                        {
                            var err = p.StandardError.ReadToEndAsync();
                            output = p.StandardOutput.ReadToEnd();
                            p.WaitForExit();
                            output += err.Result;
                        }
                        else { p.WaitForExit(); output = ""; }
                    }
                }
                catch (Exception ex) { output = "Не удалось запустить Python: " + ex.Message; }
                if (done != null && !owner.IsDisposed) owner.BeginInvoke(done, output);
            }) { IsBackground = true }.Start();
        }
    }

    // ---------- окно «Какие виды строить?» (перед построением чертежей) ----------
    class SidesForm : Form
    {
        public int Result;     // 1 — вид слева, 2 — вид справа, 3 — оба, 0 — отмена
        readonly SettingsStore Store;
        readonly string Section;
        readonly CheckBox remember;

        public SidesForm(string root, string section, string title)
        {
            Store = new SettingsStore(root);
            Section = section;
            Text = string.IsNullOrEmpty(title) ? "Чертежи рам" : title;
            Font = new Font("Segoe UI", 10f);
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = true;
            KeyPreview = true;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            Controls.Add(new Label { Text = "Какие виды строить?", Left = 16, Top = 14, AutoSize = true, Font = new Font(Font, FontStyle.Bold) });
            var current = Current();      // что выбрано в настройках — кнопка по умолчанию
            var choices = new[] { Tuple.Create("both", "Оба вида", 3), Tuple.Create("left", "Вид слева", 1), Tuple.Create("right", "Вид справа", 2) };
            int x = 16;
            foreach (var c in choices)
            {
                var code = c.Item3; var key = c.Item1;
                var b = new Button { Text = c.Item2, Left = x, Top = 46, Width = 130, Height = 40 };
                b.Click += (s, e) => Choose(code, key);
                Controls.Add(b);
                if (key == current) AcceptButton = b;
                x += 140;
            }
            remember = new CheckBox { Text = "Больше не спрашивать (поменять — в настройках менеджера)", Left = 16, Top = 98, AutoSize = true };
            Controls.Add(remember);
            var cancel = new Button { Text = "Отмена", Left = 16 + 280 + 30, Top = 128, Width = 100, DialogResult = DialogResult.Cancel };
            cancel.Click += (s, e) => { Result = 0; Close(); };
            Controls.Add(cancel);
            CancelButton = cancel;
            ClientSize = new Size(16 + 3 * 140 + 24, 166);
            cancel.Left = ClientSize.Width - 16 - cancel.Width;
            // клавиши: Л/L — слева, П/R — справа, О/B — оба
            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.L || e.KeyCode == Keys.K) Choose(1, "left");      // Л на русской раскладке = K
                else if (e.KeyCode == Keys.R || e.KeyCode == Keys.G) Choose(2, "right"); // П = G
                else if (e.KeyCode == Keys.B || e.KeyCode == Keys.J) Choose(3, "both");  // О = J
            };
            Shown += (s, e) => { Activate(); if (AcceptButton != null) ((Button)AcceptButton).Focus(); TopMost = false; };
        }

        string Current()
        {
            var it = Store.Sections.SelectMany(s => s.Items).FirstOrDefault(i => i.Section == Section && i.Key == "sides");
            return it != null ? Store.Get(it) : "both";
        }

        void Choose(int code, string key)
        {
            Result = code;
            if (remember.Checked)
            {
                foreach (var it in Store.Sections.SelectMany(s => s.Items).Where(i => i.Section == Section))
                {
                    if (it.Key == "sides") Store.Set(it, key);
                    if (it.Key == "ask_sides") Store.Set(it, "false");
                }
                try { Store.Save(); } catch { }
            }
            Close();
        }
    }

    // ---------- окно «Произошла ошибка» / «Сообщить о проблеме» ----------
    // Отправку делает _setup\report.py --send (Telegram-бот из update.ini [report]).
    class ReportForm : Form
    {
        readonly string Root, Plugin, ErrorText;
        TextBox txtComment;
        Label lblStatus;
        Button btnSend;

        public ReportForm(string root, string plugin, string errorFile, string title)
        {
            Root = root;
            Plugin = plugin ?? "";
            ErrorText = errorFile != null && File.Exists(errorFile) ? File.ReadAllText(errorFile, Encoding.UTF8) : "";
            var ver = File.Exists(Path.Combine(root, "VERSION")) ? File.ReadAllText(Path.Combine(root, "VERSION")).Trim() : "?";
            Text = (string.IsNullOrEmpty(title) ? "Плагины КОМПАС — версия " + ver : title);
            Font = new Font("Segoe UI", 9.5f);
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = true;
            ClientSize = new Size(640, ErrorText.Length > 0 ? 470 : 280);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            var y = 12;
            var head = new Label
            {
                Left = 12, Top = y, Width = 616, Height = 24, Font = new Font(Font, FontStyle.Bold),
                Text = ErrorText.Length > 0 ? "Произошла ошибка" + (Plugin.Length > 0 ? " в плагине «" + Plugin + "»" : "") + "."
                                            : "Отправить журнал работы разработчику."
            };
            Controls.Add(head);
            y += 28;
            if (ErrorText.Length > 0)
            {
                var err = new TextBox
                {
                    Left = 12, Top = y, Width = 616, Height = 180, Multiline = true, ReadOnly = true,
                    ScrollBars = ScrollBars.Both, WordWrap = false, Font = new Font("Consolas", 8.5f),
                    Text = ErrorText.Replace("\r\n", "\n").Replace("\n", "\r\n")
                };
                Controls.Add(err);
                y += 188;
            }
            Controls.Add(new Label { Left = 12, Top = y, AutoSize = true, Text = "Что делали перед этим (необязательно):" });
            y += 22;
            txtComment = new TextBox { Left = 12, Top = y, Width = 616, Height = 70, Multiline = true, ScrollBars = ScrollBars.Vertical };
            Controls.Add(txtComment);
            y += 78;
            lblStatus = new Label { Left = 12, Top = y, Width = 616, Height = 40, ForeColor = Color.DimGray,
                                    Text = "Уйдёт: описание ошибки, версия, компьютер и журнал работы плагина." };
            Controls.Add(lblStatus);
            y += 44;

            btnSend = new Button { Text = "Отправить разработчику", AutoSize = true, Top = y };
            var bLog = new Button { Text = "Открыть журнал", AutoSize = true, Top = y };
            var bClose = new Button { Text = "Закрыть", AutoSize = true, Top = y, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { btnSend, bLog, bClose });
            bClose.Left = 628 - bClose.PreferredSize.Width;
            bLog.Left = bClose.Left - 8 - bLog.PreferredSize.Width;
            btnSend.Left = bLog.Left - 8 - btnSend.PreferredSize.Width;
            CancelButton = bClose;
            ClientSize = new Size(640, y + btnSend.PreferredSize.Height + 12);

            btnSend.Click += (s, e) => Send();
            bLog.Click += (s, e) => OpenLog();
            Shown += (s, e) => { TopMost = false; Activate(); txtComment.Focus(); };
        }

        string Script() { return Path.Combine(Root, "_setup", "report.py"); }

        string PluginArg() { return Plugin.Length > 0 ? " --plugin \"" + Plugin + "\"" : ""; }

        void OpenLog()
        {
            Py.Run(this, Root, "\"" + Script() + "\" --log-path" + PluginArg(), true, output =>
            {
                var path = output.Trim().Split('\n').Last().Trim();
                if (File.Exists(path)) Py.ShowText(this, path, btnSend.Enabled ? (Action)Send : null);
                else lblStatus.Text = "Журнала пока нет: " + path;
            });
        }

        void Send()
        {
            btnSend.Enabled = false;
            lblStatus.Text = "Отправляю…";
            var tmp = Path.GetTempPath();
            var args = "\"" + Script() + "\" --send" + PluginArg();
            if (ErrorText.Length > 0)
            {
                var ef = Path.Combine(tmp, "kp_report_error.txt");
                File.WriteAllText(ef, ErrorText, new UTF8Encoding(false));
                args += " --error-file \"" + ef + "\"";
            }
            if (txtComment.Text.Trim().Length > 0)
            {
                var cf = Path.Combine(tmp, "kp_report_comment.txt");
                File.WriteAllText(cf, txtComment.Text, new UTF8Encoding(false));
                args += " --comment-file \"" + cf + "\"";
            }
            if (Text.Length > 0 && ErrorText.Length > 0) args += " --title \"ОШИБКА: " + Text.Replace("\"", "'") + "\"";
            Py.Run(this, Root, args, true, output =>
            {
                var last = output.Trim().Split('\n').Last().Trim();
                bool ok = last.StartsWith("Отчёт отправлен");
                lblStatus.ForeColor = ok ? Color.DarkGreen : Color.DarkRed;
                lblStatus.Text = ok ? last + " Спасибо!" : (last.Length > 0 ? last : "Не отправилось.") + " Можно переслать файл журнала вручную.";
                btnSend.Enabled = !ok;
            });
        }
    }
}
