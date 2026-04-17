using LiteFlow.Controller;
using LiteFlow.Core;
using LiteFlow.Forms;
using LiteFlow.Models;
using LiteFlow.Services;
using LiteTools.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LiteFlow
{
    public class LiteFlowUI : UserControl, ILitePlugin, IImageSubscriber
    {
        public string Name => "LiteFlow";
        public string Version => "1.0.0";

        private IImagePublisher? _publisher;
        private Queue<Bitmap> _pendingImages = new Queue<Bitmap>();

        private string _currentProjectPath = "";
        private LiteFlowProjectData _currentProjectData = new LiteFlowProjectData();
        private EvidenceItem? _currentEvidence;

        private string _baseDir = Path.Combine(Application.StartupPath, "LiteFlow_Data");
        private string _configPath;
        private string _templatesDir;
        private string _sessionTempDir;

        private string _defaultTemplatePath = "";
        private string _defaultQAName = "";
        private string _defaultPrefix = "";
        private bool _isRecording = true;
        private bool _isRibbonLocked = false;
        private int _savedRibbonHeight = 160;
        private bool _isAutoSaveEnabled = false;

        private LayoutMode _defaultLayoutMode = LayoutMode.Padrao;
        private int _defaultMobileColumns = 2;

        private bool _isLoadingProject = false;
        private bool _hasUnsavedChanges = false;
        private bool _firstLoadDone = false;

        private Dictionary<EvidenceItem, string> _evidenceDiskPaths = new Dictionary<EvidenceItem, string>();
        private List<EvidenceItem> _activeEvidencesList = new List<EvidenceItem>();

        private ImageEditorCore _editorCore = null!;

        private class ProjectAction
        {
            public Action UndoAction { get; set; } = null!;
            public Action RedoAction { get; set; } = null!;
        }
        private Stack<ProjectAction> _projectUndoStack = new Stack<ProjectAction>();
        private Stack<ProjectAction> _projectRedoStack = new Stack<ProjectAction>();

        private ToolStrip _topToolbar = null!;
        private ToolStripButton _btnToggleCapture = null!;
        private ToolStripButton _btnAutoSaveToggle = null!;
        private ToolStripLabel _lblProjectName = null!;

        private Panel _editorContainer = null!;
        private PictureBox _mainCanvas = null!;
        private Panel _propertiesPanel = null!;
        private FlowLayoutPanel _historyRibbon = null!;
        private SplitContainer _mainSplitter = null!;
        private Button _btnColorPicker = null!;
        private TrackBar _trkThickness = null!;
        private TextBox _floatingTextBox = null!;
        private RichTextBox _stepNoteTextBox = null!;
        private CheckBox _chkTextBelowStep = null!;
        private CheckBox _chkEvidenceOnly = null!;
        private ToolStripComboBox _cmbFont = null!;
        private ToolStripComboBox _cmbSize = null!;
        private PictureBox _templateThumbnail = null!;
        private CheckBox _chkLockRibbon = null!;

        public LiteFlowUI()
        {
            _configPath = Path.Combine(_baseDir, "settings.ini");
            _templatesDir = Path.Combine(_baseDir, "Templates");

            _sessionTempDir = Path.Combine(Path.GetTempPath(), "LiteFlowSession_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_sessionTempDir);

            // 1. Carrega APENAS o idioma antes de construir a UI (para os textos ficarem corretos)
            LoadLanguageOnly();

            // 2. Constrói os controlos da UI (botões, sliders, menus)
            InitializeComponent();

            // 3. Agora que os controlos existem, aplica as cores, espessuras e layouts guardados
            LoadSettings();

            _editorCore = new ImageEditorCore(_mainCanvas, _floatingTextBox)
            {
                CurrentColor = _btnColorPicker.BackColor,
                CurrentThickness = _trkThickness.Value,
                CurrentFontFamily = _cmbFont.Text,
                CurrentFontSize = int.Parse(_cmbSize.Text)
            };

            _editorCore.OnImageEdited = () => {
                if (_currentEvidence != null && _editorCore.WorkingImage != null)
                {
                    _currentEvidence.Image?.Dispose();
                    _currentEvidence.Image = new Bitmap(_editorCore.WorkingImage);

                    _currentEvidence.Thumbnail.Image?.Dispose();
                    _currentEvidence.Thumbnail.Image = CreateThumbnail(_currentEvidence.Image);
                    _currentEvidence.Thumbnail.Invalidate();

                    if (!string.IsNullOrEmpty(_currentEvidence.DiskPath))
                    {
                        var cloneToSave = new Bitmap(_currentEvidence.Image);
                        string path = _currentEvidence.DiskPath;
                        Task.Run(() => {
                            try { cloneToSave.Save(path, ImageFormat.Png); }
                            catch { }
                            finally { cloneToSave.Dispose(); }
                        });
                    }
                    TriggerAutoSave();
                }
            };
        }

        public void Initialize(IImagePublisher publisher, string currentLanguage)
        {
            _publisher = publisher;
            LanguageManager.CurrentLanguage = currentLanguage;

            if (_btnToggleCapture != null)
            {
                _btnToggleCapture.Enabled = true;
                _btnToggleCapture.Text = _isRecording ? LanguageManager.GetString("BtnCapturing") : LanguageManager.GetString("BtnPaused");
                _btnToggleCapture.BackColor = _isRecording ? Color.LightGreen : Color.LightCoral;
                _btnToggleCapture.ToolTipText = LanguageManager.GetString("TooltipPause");
            }

            UpdateProjectNameUI();
            UpdateAutoSaveUI();
        }

        public UserControl GetSettingsUI() { return this; }

        public void Shutdown()
        {
            ClearEvidenceHistory();
            try { if (Directory.Exists(_sessionTempDir)) Directory.Delete(_sessionTempDir, true); } catch { }
        }

        public void ReceiveImage(Bitmap image)
        {
            if (image == null || !_isRecording || this.IsDisposed) return;
            Bitmap clonedImage = new Bitmap(image);

            if (!this.IsHandleCreated)
            {
                _pendingImages.Enqueue(clonedImage);
            }
            else
            {
                Task.Run(() => {
                    string diskPath = Path.Combine(_sessionTempDir, $"img_{Guid.NewGuid():N}.png");
                    clonedImage.Save(diskPath, ImageFormat.Png);

                    this.BeginInvoke(new Action(() => {
                        if (!this.IsDisposed) AddToHistoryFromDisk(diskPath, clonedImage);
                        else clonedImage.Dispose();
                    }));
                });
            }
        }

        public void ReceiveImage(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath) || !_isRecording || this.IsDisposed) return;
            this.BeginInvoke(new Action(() => {
                if (!this.IsDisposed) AddToHistoryFromDisk(imagePath, null);
            }));
        }

        private void AddToHistoryFromDisk(string diskPath, Bitmap? alreadyLoadedImage = null, string note = "", bool textBelow = false, bool isEvidenceOnly = false)
        {
            if (!File.Exists(diskPath)) return;

            Bitmap activeImg = alreadyLoadedImage ?? LoadImageFromDisk(diskPath);
            Bitmap thumbImg = CreateThumbnail(activeImg);

            var item = new EvidenceItem { Image = activeImg, Note = note, TextBelowImage = textBelow, IsEvidenceOnly = isEvidenceOnly, DiskPath = diskPath };

            var thumbnailPb = new PictureBox { Width = 140, Height = 90, SizeMode = PictureBoxSizeMode.Zoom, Image = thumbImg, BackColor = Color.White, Padding = new Padding(2), Margin = new Padding(5), Cursor = Cursors.Hand, Tag = item };
            item.Thumbnail = thumbnailPb;
            thumbnailPb.Paint += Thumbnail_Paint;
            thumbnailPb.MouseDown += Thumbnail_MouseDown;

            InsertEvidenceToUI(item, _historyRibbon.Controls.Count, false);
            if (_stepNoteTextBox.Focused) TriggerAutoSave();

            SelectEvidence(item);
        }

        private Bitmap LoadImageFromDisk(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var img = Image.FromStream(fs))
            {
                return new Bitmap(img);
            }
        }

        private Bitmap CreateThumbnail(Bitmap original)
        {
            int w = 140;
            int h = (int)((140.0f / original.Width) * original.Height);
            if (h <= 0) h = 90;

            var thumb = new Bitmap(w, h, PixelFormat.Format16bppRgb555);
            using (var g = Graphics.FromImage(thumb))
            {
                g.InterpolationMode = InterpolationMode.Low;
                g.SmoothingMode = SmoothingMode.HighSpeed;
                g.CompositingQuality = CompositingQuality.HighSpeed;
                g.DrawImage(original, 0, 0, w, h);
            }
            return thumb;
        }

        private void EnforceRuleOfThree(EvidenceItem newlyActivated)
        {
            if (!_activeEvidencesList.Contains(newlyActivated))
                _activeEvidencesList.Add(newlyActivated);
            else
            {
                _activeEvidencesList.Remove(newlyActivated);
                _activeEvidencesList.Add(newlyActivated);
            }

            while (_activeEvidencesList.Count > 3)
            {
                var oldestItem = _activeEvidencesList[0];
                _activeEvidencesList.RemoveAt(0);

                if (oldestItem != newlyActivated && oldestItem != _currentEvidence)
                {
                    if (oldestItem.Image != null)
                    {
                        oldestItem.Image.Dispose();
                        oldestItem.Image = null!;
                    }
                }
            }
            GC.Collect(2, GCCollectionMode.Forced, true);
        }

        private void SelectEvidence(EvidenceItem? item)
        {
            if (_currentEvidence != null && _currentEvidence.Thumbnail != null) _currentEvidence.Thumbnail.BackColor = Color.White;
            _currentEvidence = item;

            if (_currentEvidence == null)
            {
                _editorCore?.LoadImage(new Bitmap(10, 10));
                _stepNoteTextBox.Text = ""; _stepNoteTextBox.Enabled = false;
                _chkTextBelowStep.Checked = false; _chkTextBelowStep.Enabled = false;
                if (_chkEvidenceOnly != null) { _chkEvidenceOnly.Checked = false; _chkEvidenceOnly.Enabled = false; }
                return;
            }

            _currentEvidence.Thumbnail.BackColor = Color.FromArgb(0, 120, 215);

            if (_currentEvidence.Image == null)
            {
                if (!string.IsNullOrEmpty(_currentEvidence.DiskPath) && File.Exists(_currentEvidence.DiskPath))
                    _currentEvidence.Image = LoadImageFromDisk(_currentEvidence.DiskPath);
                else
                    _currentEvidence.Image = new Bitmap(1024, 768);
            }

            EnforceRuleOfThree(_currentEvidence);

            _editorCore?.LoadImage(_currentEvidence.Image);
            _stepNoteTextBox.Enabled = true; _stepNoteTextBox.Text = _currentEvidence.Note;

            _chkTextBelowStep.Enabled = true; _chkTextBelowStep.Checked = _currentEvidence.TextBelowImage;
            _chkEvidenceOnly.Enabled = true; _chkEvidenceOnly.Checked = _currentEvidence.IsEvidenceOnly;

            foreach (ToolStripItem t in _topToolbar.Items)
                if (t is ToolStripButton b && b.Tag is EditorTool et && et == EditorTool.Select) { b.PerformClick(); break; }
        }

        private void SaveProjectInternal(string path)
        {
            try
            {
                using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var writer = new System.Text.Json.Utf8JsonWriter(fileStream))
                {
                    writer.WriteStartObject();
                    writer.WriteString("TemplatePath", _currentProjectData.TemplatePath ?? "");
                    writer.WriteString("FilePrefix", _currentProjectData.FilePrefix ?? "");
                    writer.WriteString("FileName", _currentProjectData.FileName ?? "");
                    writer.WriteString("TestCaseName", _currentProjectData.TestCaseName ?? "");
                    writer.WriteString("QAName", _currentProjectData.QAName ?? "");
                    writer.WriteString("TestDate", _currentProjectData.TestDate ?? "");
                    writer.WriteString("Comments", _currentProjectData.Comments ?? "");
                    writer.WriteNumber("ReportLayout", (int)_currentProjectData.ReportLayout);
                    writer.WriteNumber("MobileColumns", _currentProjectData.MobileColumns);

                    writer.WritePropertyName("Steps");
                    writer.WriteStartArray();

                    var items = GetItems();
                    foreach (var item in items)
                    {
                        writer.WriteStartObject();
                        string base64 = "";

                        if (!string.IsNullOrEmpty(item.DiskPath) && File.Exists(item.DiskPath))
                        {
                            base64 = Convert.ToBase64String(File.ReadAllBytes(item.DiskPath));
                        }
                        else if (item.Image != null)
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                item.Image.Save(ms, ImageFormat.Png);
                                base64 = Convert.ToBase64String(ms.ToArray());
                            }
                        }

                        writer.WriteString("ImageDataBase64", base64);
                        writer.WriteString("Note", item.Note ?? "");
                        writer.WriteBoolean("TextBelowImage", item.TextBelowImage);
                        writer.WriteBoolean("IsEvidenceOnly", item.IsEvidenceOnly);
                        writer.WriteEndObject();

                        base64 = null!;
                    }

                    writer.WriteEndArray();
                    writer.WriteEndObject();
                }
            }
            catch { }

            GC.Collect(2, GCCollectionMode.Forced, true);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            this.DoubleBuffered = true;

            SetupToolbar();

            _mainSplitter = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, FixedPanel = FixedPanel.Panel2, SplitterWidth = 4 };
            SetupEditorArea(_mainSplitter.Panel1);
            SetupHistoryRibbon(_mainSplitter.Panel2);

            _mainSplitter.SplitterMoved += (s, e) => {
                if (!_isLoadingProject && _firstLoadDone)
                {
                    _savedRibbonHeight = _mainSplitter.Height - _mainSplitter.SplitterDistance;
                    SaveSettings();
                }
            };

            this.Controls.Add(_mainSplitter);
            this.Controls.Add(_topToolbar);
            this.ResumeLayout(false);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            while (_pendingImages.Count > 0)
            {
                Bitmap pendingImg = _pendingImages.Dequeue();
                Task.Run(() => {
                    string diskPath = Path.Combine(_sessionTempDir, $"img_{Guid.NewGuid():N}.png");
                    pendingImg.Save(diskPath, ImageFormat.Png);
                    this.BeginInvoke(new Action(() => {
                        if (!this.IsDisposed) AddToHistoryFromDisk(diskPath, pendingImg);
                        else pendingImg.Dispose();
                    }));
                });
            }

            this.BeginInvoke(new Action(() => {
                if (this.IsDisposed) return;

                int safeDistance = _mainSplitter.Height - _savedRibbonHeight;
                if (safeDistance > 50 && safeDistance < _mainSplitter.Height - 50)
                    _mainSplitter.SplitterDistance = safeDistance;
                else
                    _mainSplitter.SplitterDistance = Math.Max(this.Height - 160, 50);

                _mainSplitter.IsSplitterFixed = _isRibbonLocked;
                if (_chkLockRibbon != null)
                {
                    _chkLockRibbon.Checked = _isRibbonLocked;
                    _chkLockRibbon.BackColor = _isRibbonLocked ? Color.LightGray : Color.White;
                }

                if (!_firstLoadDone)
                {
                    _firstLoadDone = true;
                    if (string.IsNullOrEmpty(_currentProjectPath) && _historyRibbon.Controls.Count <= 1)
                    {
                        var res = MessageBox.Show(LanguageManager.GetString("MsgStartNewForAutoSave"), LanguageManager.GetString("TitleNewProject"), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (res == DialogResult.Yes) { SaveProjectAs(); }
                    }
                }
            }));
        }

        private void LoadLanguageOnly()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    var lines = File.ReadAllLines(_configPath);
                    if (lines.Length > 12 && !string.IsNullOrEmpty(lines[12]))
                    {
                        LanguageManager.CurrentLanguage = lines[12];
                    }
                }
            }
            catch { }
        }

        private void LoadSettings()
        {
            try
            {
                Directory.CreateDirectory(_baseDir);
                if (File.Exists(_configPath))
                {
                    var lines = File.ReadAllLines(_configPath);
                    if (lines.Length > 0 && int.TryParse(lines[0], out int rHeight)) _savedRibbonHeight = rHeight < 400 ? rHeight : 160;
                    if (lines.Length > 1) { _btnColorPicker.BackColor = ColorTranslator.FromHtml(lines[1]); }
                    if (lines.Length > 2 && int.TryParse(lines[2], out int thick)) { _trkThickness.Value = Math.Max(1, Math.Min(20, thick)); }
                    if (lines.Length > 3) _cmbFont.Text = lines[3];
                    if (lines.Length > 4) _cmbSize.Text = lines[4];
                    if (lines.Length > 5) _defaultTemplatePath = lines[5];
                    if (lines.Length > 6) _defaultQAName = lines[6];
                    if (lines.Length > 7) _defaultPrefix = lines[7];
                    if (lines.Length > 8 && bool.TryParse(lines[8], out bool isLocked)) _isRibbonLocked = isLocked;
                    if (lines.Length > 10 && Enum.TryParse(lines[10], out LayoutMode lMode)) _defaultLayoutMode = lMode;
                    if (lines.Length > 11 && int.TryParse(lines[11], out int cols)) _defaultMobileColumns = cols;
                    // Idioma já foi carregado no LoadLanguageOnly, logo ignoramos a linha 12 aqui.
                }
            }
            catch { }
        }

        private void SaveSettings()
        {
            try
            {
                Directory.CreateDirectory(_baseDir);
                var lines = new string[] {
                    _savedRibbonHeight.ToString(), ColorTranslator.ToHtml(_btnColorPicker.BackColor),
                    _trkThickness.Value.ToString(), _cmbFont.Text, _cmbSize.Text,
                    _defaultTemplatePath, _defaultQAName, _defaultPrefix, _isRibbonLocked.ToString(),
                    "false",
                    _defaultLayoutMode.ToString(),
                    _defaultMobileColumns.ToString(),
                    LanguageManager.CurrentLanguage
                };
                File.WriteAllLines(_configPath, lines);
                _templateThumbnail?.Invalidate();
            }
            catch { }
        }

        private void PerformUndo()
        {
            if (_editorCore != null && _editorCore.CanUndo) { _editorCore.Undo(); }
            else if (_projectUndoStack.Count > 0)
            {
                var action = _projectUndoStack.Pop();
                action.UndoAction();
                _projectRedoStack.Push(action);
            }
        }

        private void PerformRedo()
        {
            if (_editorCore != null && _editorCore.CanRedo) { _editorCore.Redo(); }
            else if (_projectRedoStack.Count > 0)
            {
                var action = _projectRedoStack.Pop();
                action.RedoAction();
                _projectUndoStack.Push(action);
            }
        }

        private void SetupToolbar()
        {
            _topToolbar = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, BackColor = Color.White, RenderMode = ToolStripRenderMode.Professional, Padding = new Padding(5) };

            _btnToggleCapture = new ToolStripButton(LanguageManager.GetString("BtnCapturing")) { BackColor = Color.LightGray, Enabled = false, ToolTipText = LanguageManager.GetString("TooltipStandalone") };
            _btnToggleCapture.Click += (s, e) => {
                _isRecording = !_isRecording;
                _btnToggleCapture.Text = _isRecording ? LanguageManager.GetString("BtnCapturing") : LanguageManager.GetString("BtnPaused");
                _btnToggleCapture.BackColor = _isRecording ? Color.LightGreen : Color.LightCoral;
            };

            var btnNew = new ToolStripButton(LanguageManager.GetString("BtnNew")) { ToolTipText = LanguageManager.GetString("TooltipNew") }; btnNew.Click += (s, e) => NewProject();
            var btnOpen = new ToolStripButton(LanguageManager.GetString("BtnOpen")) { ToolTipText = LanguageManager.GetString("TooltipOpen") }; btnOpen.Click += (s, e) => OpenProject();
            var btnSave = new ToolStripButton(LanguageManager.GetString("BtnSave")) { ToolTipText = LanguageManager.GetString("TooltipSave") }; btnSave.Click += (s, e) => SaveProjectCurrent();
            var btnSaveAs = new ToolStripButton(LanguageManager.GetString("BtnSaveAs")) { ToolTipText = LanguageManager.GetString("TooltipSaveAs") }; btnSaveAs.Click += (s, e) => SaveProjectAs();

            _btnAutoSaveToggle = new ToolStripButton(LanguageManager.GetString("AutoSaveOff")) { ForeColor = Color.Gray, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ToolTipText = LanguageManager.GetString("TooltipAutoSave") };
            _btnAutoSaveToggle.Click += (s, e) => {
                if (string.IsNullOrEmpty(_currentProjectPath))
                {
                    MessageBox.Show(LanguageManager.GetString("MsgSaveFirst"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _isAutoSaveEnabled = !_isAutoSaveEnabled;
                UpdateAutoSaveUI();
                if (_isAutoSaveEnabled) TriggerAutoSave();
            };

            var btnUndo = new ToolStripButton("↩️") { ToolTipText = "Ctrl+Z" }; btnUndo.Click += (s, e) => PerformUndo();
            var btnRedo = new ToolStripButton("↪️") { ToolTipText = "Ctrl+Y" }; btnRedo.Click += (s, e) => PerformRedo();

            _cmbFont = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };
            _cmbFont.Items.AddRange(new object[] { "Segoe UI", "Arial", "Consolas", "Times New Roman", "Verdana" });
            _cmbFont.SelectedItem = "Segoe UI";
            _cmbFont.SelectedIndexChanged += (s, e) => { if (_editorCore != null) { _editorCore.CurrentFontFamily = _cmbFont.Text; SaveSettings(); } };

            _cmbSize = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 50 };
            _cmbSize.Items.AddRange(new object[] { "10", "12", "14", "16", "18", "24", "32", "48", "72", "96" });
            _cmbSize.SelectedItem = "14";
            _cmbSize.SelectedIndexChanged += (s, e) => { if (_editorCore != null && int.TryParse(_cmbSize.Text, out int size)) { _editorCore.CurrentFontSize = size; SaveSettings(); } };

            var btnGlobalSettings = new ToolStripButton("⚙️") { ToolTipText = LanguageManager.GetString("SettingsTitle"), Alignment = ToolStripItemAlignment.Right };
            btnGlobalSettings.Click += BtnGlobalSettings_Click;

            _lblProjectName = new ToolStripLabel(LanguageManager.GetString("ProjectNotSaved")) { ForeColor = Color.DimGray, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Alignment = ToolStripItemAlignment.Right };

            var tools = new ToolStripItem[] {
                _btnToggleCapture, new ToolStripSeparator(),
                btnNew, btnOpen, btnSave, btnSaveAs, _btnAutoSaveToggle, new ToolStripSeparator(),
                btnUndo, btnRedo, new ToolStripSeparator(),
                new ToolStripButton(LanguageManager.GetString("ToolSelect")) { Tag = EditorTool.Select, ToolTipText = LanguageManager.GetString("TooltipSelect") },
                new ToolStripButton(LanguageManager.GetString("ToolCrop")) { Tag = EditorTool.Crop, ToolTipText = LanguageManager.GetString("TooltipCrop") },
                new ToolStripSeparator(),
                new ToolStripButton(LanguageManager.GetString("ToolPen")) { Tag = EditorTool.Pen },
                new ToolStripButton(LanguageManager.GetString("ToolLine")) { Tag = EditorTool.Line },
                new ToolStripButton(LanguageManager.GetString("ToolArrow")) { Tag = EditorTool.Arrow, Checked = true },
                new ToolStripButton(LanguageManager.GetString("ToolShape")) { Tag = EditorTool.Shape },
                new ToolStripButton(LanguageManager.GetString("ToolHighlight")) { Tag = EditorTool.Highlight },
                new ToolStripButton(LanguageManager.GetString("ToolText")) { Tag = EditorTool.Text },
                new ToolStripSeparator(),
                new ToolStripLabel(LanguageManager.GetString("LblFont")), _cmbFont, _cmbSize,
                btnGlobalSettings,
                _lblProjectName
            };

            foreach (var item in tools)
            {
                if (item is ToolStripButton btn && btn.Tag is EditorTool)
                {
                    btn.Click += (s, e) =>
                    {
                        foreach (ToolStripItem t in _topToolbar.Items) if (t is ToolStripButton b && b.Tag is EditorTool) b.Checked = false;
                        btn.Checked = true;
                        _editorCore.CurrentTool = (EditorTool)btn.Tag;
                        _mainCanvas.Cursor = (_editorCore.CurrentTool == EditorTool.Text) ? Cursors.IBeam : Cursors.Cross;
                        _editorCore.CancelCurrentAction();
                    };
                }
            }
            _topToolbar.Items.AddRange(tools);
        }

        private void UpdateAutoSaveUI()
        {
            if (string.IsNullOrEmpty(_currentProjectPath))
            {
                _isAutoSaveEnabled = false;
                _btnAutoSaveToggle.Text = LanguageManager.GetString("AutoSaveOff");
                _btnAutoSaveToggle.ForeColor = Color.Gray;
            }
            else
            {
                _btnAutoSaveToggle.Text = _isAutoSaveEnabled ? LanguageManager.GetString("AutoSaveOn") : LanguageManager.GetString("AutoSaveOff");
                _btnAutoSaveToggle.ForeColor = _isAutoSaveEnabled ? Color.Green : Color.Orange;
            }
        }

        private void UpdateProjectNameUI()
        {
            if (string.IsNullOrEmpty(_currentProjectPath)) _lblProjectName.Text = LanguageManager.GetString("ProjectNotSaved");
            else _lblProjectName.Text = string.Format(LanguageManager.GetString("ProjectName"), Path.GetFileName(_currentProjectPath));
        }

        private void BtnGlobalSettings_Click(object? sender, EventArgs e)
        {
            using (Form frm = new Form())
            {
                frm.Text = LanguageManager.GetString("SettingsTitle");
                frm.Size = new Size(400, 260);
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.FormBorderStyle = FormBorderStyle.FixedDialog;
                frm.MaximizeBox = false;
                frm.MinimizeBox = false;
                frm.BackColor = Color.White;
                frm.Font = new Font("Segoe UI", 9F);

                var pbIcon = new PictureBox { Image = SystemIcons.Information.ToBitmap(), Location = new Point(20, 20), Size = new Size(32, 32), SizeMode = PictureBoxSizeMode.Zoom };

                var lblTitle = new Label { Text = LanguageManager.GetString("SettingsHeader"), Font = new Font("Segoe UI", 12F, FontStyle.Bold), Location = new Point(60, 20), AutoSize = true, ForeColor = Color.FromArgb(0, 120, 215) };
                var lblVersion = new Label { Text = $"{LanguageManager.GetString("SettingsVersion")}{this.Version}", Location = new Point(62, 45), AutoSize = true, ForeColor = Color.DimGray };

                var lblLang = new Label { Text = LanguageManager.GetString("SettingsLanguage"), Location = new Point(62, 80), AutoSize = true };
                var cmbLang = new ComboBox { Location = new Point(62, 100), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
                cmbLang.Items.AddRange(new string[] { "pt-BR", "en-US", "es-ES", "fr-FR", "de-DE", "it-IT" });

                if (cmbLang.Items.Contains(LanguageManager.CurrentLanguage)) cmbLang.SelectedItem = LanguageManager.CurrentLanguage;
                else cmbLang.SelectedIndex = 0;

                bool isPluginMode = _publisher != null;
                if (isPluginMode)
                {
                    cmbLang.Enabled = false;
                    var lblInfo = new Label { Text = LanguageManager.GetString("LangManagedByHost"), Location = new Point(220, 102), AutoSize = true, ForeColor = Color.Gray };
                    frm.Controls.Add(lblInfo);
                }

                var lnkGithub = new LinkLabel { Text = "GitHub: github.com/eugenio122/LiteFlow", Location = new Point(62, 140), AutoSize = true, LinkColor = Color.FromArgb(0, 120, 215) };
                lnkGithub.LinkClicked += (s, ev) => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://github.com/eugenio122/LiteFlow") { UseShellExecute = true });

                var btnOk = new Button { Text = LanguageManager.GetString("BtnSaveClose"), Location = new Point(140, 180), Width = 120, Height = 30, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
                btnOk.FlatAppearance.BorderSize = 0;
                btnOk.Click += (s, ev) => {
                    bool changed = false;
                    if (!isPluginMode)
                    {
                        var sel = cmbLang.SelectedItem?.ToString() ?? "pt-BR";
                        if (sel != LanguageManager.CurrentLanguage)
                        {
                            LanguageManager.CurrentLanguage = sel;
                            changed = true;
                        }
                    }
                    SaveSettings();
                    frm.DialogResult = DialogResult.OK;
                    frm.Close();

                    if (changed) MessageBox.Show(LanguageManager.GetString("MsgRestartRequired"), "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                frm.Controls.AddRange(new Control[] { pbIcon, lblTitle, lblVersion, lblLang, cmbLang, lnkGithub, btnOk });
                frm.ShowDialog();
            }
        }

        private void SetupEditorArea(Control parent)
        {
            _editorContainer = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(220, 220, 220) };
            SetupPropertiesPanel(_editorContainer);

            Panel imageContainer = new Panel { Dock = DockStyle.Fill };
            Panel notePanel = new Panel { Dock = DockStyle.Top, Height = 65, Padding = new Padding(5), BackColor = Color.WhiteSmoke };

            Panel noteHeaderPanel = new Panel { Dock = DockStyle.Top, Height = 25 };
            Label lblNote = new Label { Text = LanguageManager.GetString("EvidenceNotes"), Dock = DockStyle.Left, AutoSize = true, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), ForeColor = Color.DimGray, Padding = new Padding(0, 4, 0, 0) };

            _chkTextBelowStep = new CheckBox { Text = LanguageManager.GetString("TextBelowImage"), Dock = DockStyle.Right, AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.DimGray, Cursor = Cursors.Hand, Enabled = false };
            _chkTextBelowStep.CheckedChanged += (s, e) => {
                if (_currentEvidence != null && _currentEvidence.TextBelowImage != _chkTextBelowStep.Checked)
                {
                    _currentEvidence.TextBelowImage = _chkTextBelowStep.Checked;
                    _hasUnsavedChanges = true;
                    TriggerAutoSave();
                }
            };

            _chkEvidenceOnly = new CheckBox { Text = $"👁️ {LanguageManager.GetString("EvidenceOnly")}", Dock = DockStyle.Right, AutoSize = true, Font = new Font("Segoe UI", 8F), ForeColor = Color.DimGray, Cursor = Cursors.Hand, Enabled = false };
            _chkEvidenceOnly.CheckedChanged += (s, e) => {
                if (_currentEvidence != null && _currentEvidence.IsEvidenceOnly != _chkEvidenceOnly.Checked)
                {
                    _currentEvidence.IsEvidenceOnly = _chkEvidenceOnly.Checked;
                    _currentEvidence.Thumbnail.Invalidate();
                    _hasUnsavedChanges = true;
                    TriggerAutoSave();
                }
            };

            noteHeaderPanel.Controls.Add(_chkEvidenceOnly);
            noteHeaderPanel.Controls.Add(_chkTextBelowStep);
            noteHeaderPanel.Controls.Add(lblNote);

            _stepNoteTextBox = new RichTextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = RichTextBoxScrollBars.Vertical, Enabled = false, BorderStyle = BorderStyle.None };

            var ctxMenu = new ContextMenuStrip();
            ctxMenu.Items.Add(LanguageManager.GetString("CtxPasteFormat"), null, (s, e) => _stepNoteTextBox.Paste());
            ctxMenu.Items.Add(LanguageManager.GetString("CtxPasteNoFormat"), null, (s, e) => {
                if (Clipboard.ContainsText()) _stepNoteTextBox.SelectedText = Clipboard.GetText(TextDataFormat.Text);
            });
            _stepNoteTextBox.ContextMenuStrip = ctxMenu;

            _stepNoteTextBox.TextChanged += (s, e) => {
                if (_currentEvidence != null)
                {
                    _currentEvidence.Note = _stepNoteTextBox.Text;
                    _hasUnsavedChanges = true;
                }
            };

            _stepNoteTextBox.Leave += (s, e) => { TriggerAutoSave(); };

            notePanel.Controls.Add(_stepNoteTextBox); notePanel.Controls.Add(noteHeaderPanel);

            _mainCanvas = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(200, 200, 200), Cursor = Cursors.Cross };
            _floatingTextBox = new TextBox { Visible = false, Multiline = true, BorderStyle = BorderStyle.FixedSingle };
            _mainCanvas.Controls.Add(_floatingTextBox);

            _mainCanvas.MouseDown += (s, e) => {
                if (_stepNoteTextBox.Focused)
                {
                    _editorContainer.Focus();
                }
            };

            imageContainer.Controls.Add(_mainCanvas);
            imageContainer.Controls.Add(notePanel);

            _editorContainer.Controls.Add(imageContainer);
            imageContainer.BringToFront();
            parent.Controls.Add(_editorContainer);
        }

        public static void ForcePlainTextPaste(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                if (Clipboard.ContainsText())
                {
                    TextBox tb = (TextBox)sender!;
                    string plainText = Clipboard.GetText(TextDataFormat.Text);
                    if (!tb.Multiline) plainText = plainText.Replace("\r", "").Replace("\n", " ");

                    int selectionStart = tb.SelectionStart;
                    tb.Text = tb.Text.Remove(selectionStart, tb.SelectionLength).Insert(selectionStart, plainText);
                    tb.SelectionStart = selectionStart + plainText.Length;
                }
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        }

        private void SetupPropertiesPanel(Control parent)
        {
            _propertiesPanel = new Panel { Dock = DockStyle.Right, Width = 260, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Padding = new Padding(15) };

            var pnlColor = new Panel { Dock = DockStyle.Top, Height = 60 };
            _btnColorPicker = new Button { Dock = DockStyle.Top, Height = 30, BackColor = Color.FromArgb(0, 178, 89), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            _btnColorPicker.Click += (s, e) => {
                using (ColorDialog cd = new ColorDialog { Color = _btnColorPicker.BackColor })
                {
                    if (cd.ShowDialog() == DialogResult.OK)
                    {
                        _btnColorPicker.BackColor = cd.Color;
                        if (_editorCore != null) { _editorCore.CurrentColor = cd.Color; _floatingTextBox.ForeColor = cd.Color; SaveSettings(); }
                    }
                }
            };
            pnlColor.Controls.Add(_btnColorPicker); pnlColor.Controls.Add(new Label { Text = LanguageManager.GetString("LblColor"), Dock = DockStyle.Top, Height = 20 });

            var pnlThickness = new Panel { Dock = DockStyle.Top, Height = 70, Padding = new Padding(0, 10, 0, 0) };
            _trkThickness = new TrackBar { Dock = DockStyle.Top, Minimum = 1, Maximum = 20, Value = 4, TickStyle = TickStyle.None, Height = 30 };
            _trkThickness.Scroll += (s, e) => { if (_editorCore != null) { _editorCore.CurrentThickness = _trkThickness.Value; SaveSettings(); } };
            pnlThickness.Controls.Add(_trkThickness); pnlThickness.Controls.Add(new Label { Text = LanguageManager.GetString("LblThickness"), Dock = DockStyle.Top, Height = 20 });

            var pnlExport = new Panel { Dock = DockStyle.Bottom, Height = 145 };
            var btnSetTemplate = new Button { Text = LanguageManager.GetString("BtnImportTemplate"), Dock = DockStyle.Top, Height = 35, BackColor = Color.WhiteSmoke, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnSetTemplate.FlatAppearance.BorderColor = Color.Silver; btnSetTemplate.Click += BtnImportWord_Click;

            var btnWordExport = new Button { Text = LanguageManager.GetString("BtnExportWord"), Dock = DockStyle.Top, Height = 45, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnWordExport.FlatAppearance.BorderSize = 0; btnWordExport.Click += BtnExportWord_Click;

            var btnPdf = new Button { Text = LanguageManager.GetString("BtnExportPdf"), Dock = DockStyle.Top, Height = 45, BackColor = Color.FromArgb(220, 53, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnPdf.FlatAppearance.BorderSize = 0; btnPdf.Click += BtnApplyToPdf_Click;

            pnlExport.Controls.Add(btnPdf); pnlExport.Controls.Add(new Label { Dock = DockStyle.Top, Height = 10 });
            pnlExport.Controls.Add(btnWordExport); pnlExport.Controls.Add(new Label { Dock = DockStyle.Top, Height = 10 });
            pnlExport.Controls.Add(btnSetTemplate);

            _propertiesPanel.Controls.Add(pnlThickness); _propertiesPanel.Controls.Add(pnlColor);
            _propertiesPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblProperties"), Dock = DockStyle.Top, Font = new Font("Segoe UI", 10F, FontStyle.Bold), Height = 40, TextAlign = ContentAlignment.MiddleLeft });
            _propertiesPanel.Controls.Add(pnlExport);
            parent.Controls.Add(_propertiesPanel);
        }

        private void SetupHistoryRibbon(Control parent)
        {
            var historyWrapper = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(5) };

            Panel headerPanel = new Panel { Dock = DockStyle.Top, Height = 28, Padding = new Padding(0, 0, 0, 3) };
            Label lblTitle = new Label { Text = LanguageManager.GetString("LblHistory"), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.DimGray, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold) };

            FlowLayoutPanel pnlRightButtons = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };

            _chkLockRibbon = new CheckBox { Appearance = Appearance.Button, Text = LanguageManager.GetString("BtnPin"), Width = 60, Height = 25, FlatStyle = FlatStyle.Flat, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 8F), Cursor = Cursors.Hand };
            _chkLockRibbon.FlatAppearance.BorderSize = 0;
            _chkLockRibbon.Click += (s, e) => {
                _isRibbonLocked = _chkLockRibbon.Checked;
                _mainSplitter.IsSplitterFixed = _isRibbonLocked;
                _chkLockRibbon.BackColor = _isRibbonLocked ? Color.LightGray : Color.White;
                SaveSettings();
            };

            Button btnAddBlank = new Button { Text = LanguageManager.GetString("BtnAddBlank"), Width = 110, Height = 25, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8F), Cursor = Cursors.Hand, BackColor = Color.WhiteSmoke };
            btnAddBlank.FlatAppearance.BorderSize = 0;
            btnAddBlank.Click += (s, e) => {
                Bitmap blank = new Bitmap(1024, 768);
                using (Graphics g = Graphics.FromImage(blank)) { g.Clear(Color.White); }

                Task.Run(() => {
                    string path = Path.Combine(_sessionTempDir, $"img_{Guid.NewGuid():N}.png");
                    blank.Save(path, ImageFormat.Png);
                    this.BeginInvoke(new Action(() => { AddToHistoryFromDisk(path, blank); }));
                });
            };

            Button btnPaste = new Button { Text = LanguageManager.GetString("BtnPasteImage"), Width = 110, Height = 25, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8F), Cursor = Cursors.Hand, BackColor = Color.WhiteSmoke };
            btnPaste.FlatAppearance.BorderSize = 0;
            btnPaste.Click += (s, e) => {
                if (!PasteImageFromClipboard()) MessageBox.Show(LanguageManager.GetString("MsgNoImageClipboard"), LanguageManager.GetString("TitlePaste"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            pnlRightButtons.Controls.Add(_chkLockRibbon);
            pnlRightButtons.Controls.Add(btnAddBlank);
            pnlRightButtons.Controls.Add(btnPaste);

            headerPanel.Controls.Add(pnlRightButtons);
            headerPanel.Controls.Add(lblTitle);

            _historyRibbon = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, WrapContents = false, BackColor = Color.FromArgb(245, 245, 245), AllowDrop = true };
            _historyRibbon.DragEnter += (s, e) => { if (e.Data.GetDataPresent(typeof(PictureBox))) e.Effect = DragDropEffects.Move; };
            _historyRibbon.DragDrop += HistoryRibbon_DragDrop;

            _templateThumbnail = new PictureBox { Width = 140, Height = 90, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.WhiteSmoke, Padding = new Padding(2), Margin = new Padding(5), Cursor = Cursors.Hand };
            _templateThumbnail.Paint += TemplateThumbnail_Paint;
            _templateThumbnail.Click += (s, e) => {
                string tPath = !string.IsNullOrEmpty(_currentProjectData.TemplatePath) ? _currentProjectData.TemplatePath : _defaultTemplatePath;
                if (File.Exists(tPath)) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(tPath) { UseShellExecute = true });
                else MessageBox.Show(LanguageManager.GetString("MsgNoTemplate"), LanguageManager.GetString("TitleTemplate"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            _historyRibbon.Controls.Add(_templateThumbnail);

            historyWrapper.Controls.Add(_historyRibbon);
            historyWrapper.Controls.Add(headerPanel);
            parent.Controls.Add(historyWrapper);
        }

        private void TemplateThumbnail_Paint(object? sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.DrawRectangle(Pens.LightGray, 0, 0, _templateThumbnail.Width - 1, _templateThumbnail.Height - 1);

            string tPath = !string.IsNullOrEmpty(_currentProjectData.TemplatePath) ? _currentProjectData.TemplatePath : _defaultTemplatePath;
            bool hasTemplate = File.Exists(tPath);

            e.Graphics.FillRectangle(hasTemplate ? Brushes.White : Brushes.WhiteSmoke, 45, 10, 50, 60);
            e.Graphics.DrawRectangle(Pens.SteelBlue, 45, 10, 50, 60);
            e.Graphics.FillRectangle(Brushes.SteelBlue, 45, 10, 50, 15);

            using (Font f = new Font("Segoe UI", 7, FontStyle.Bold)) { e.Graphics.DrawString("WORD", f, Brushes.White, 53, 11); }
            using (Font f2 = new Font("Segoe UI", 7, FontStyle.Regular))
            {
                string name = hasTemplate ? Path.GetFileName(tPath) : LanguageManager.GetString("BlankDoc");
                if (name.Length > 20) name = name.Substring(0, 17) + "...";
                SizeF s = e.Graphics.MeasureString(name, f2);
                e.Graphics.DrawString(name, f2, Brushes.Black, (_templateThumbnail.Width - s.Width) / 2, 75);
            }
        }

        private void ClearEvidenceHistory()
        {
            for (int i = _historyRibbon.Controls.Count - 1; i >= 0; i--)
            {
                var ctrl = _historyRibbon.Controls[i];
                if (ctrl != _templateThumbnail)
                {
                    _historyRibbon.Controls.RemoveAt(i);

                    if (ctrl is PictureBox pb)
                    {
                        if (pb.Tag is EvidenceItem item)
                        {
                            item.Image?.Dispose();
                            item.Image = null!;
                        }
                        pb.Image?.Dispose();
                        pb.Image = null;
                    }
                    ctrl.Dispose();
                }
            }

            _projectUndoStack.Clear();
            _projectRedoStack.Clear();
            _activeEvidencesList.Clear();
            _evidenceDiskPaths.Clear();

            if (_editorCore != null) { _editorCore.LoadImage(new Bitmap(10, 10)); }

            while (_pendingImages.Count > 0)
            {
                _pendingImages.Dequeue()?.Dispose();
            }

            SelectEvidence(null);

            try
            {
                if (Directory.Exists(_sessionTempDir)) Directory.Delete(_sessionTempDir, true);
                Directory.CreateDirectory(_sessionTempDir);
            }
            catch { }

            GC.Collect(2, GCCollectionMode.Forced, true);
        }

        private void InsertEvidenceToUI(EvidenceItem item, int index, bool isUndoRedo)
        {
            _historyRibbon.Controls.Add(item.Thumbnail);
            if (index < _historyRibbon.Controls.Count)
            {
                _historyRibbon.Controls.SetChildIndex(item.Thumbnail, index);
            }

            if (!isUndoRedo && !_isLoadingProject)
            {
                _projectUndoStack.Push(new ProjectAction
                {
                    UndoAction = () => RemoveEvidenceFromUI(item, true),
                    RedoAction = () => InsertEvidenceToUI(item, index, true)
                });
                _projectRedoStack.Clear();
            }

            if (!_isLoadingProject) _historyRibbon.ScrollControlIntoView(item.Thumbnail);
            ReindexHistory(); TriggerAutoSave();
        }

        private void RemoveEvidenceFromUI(EvidenceItem item, bool isUndoRedo)
        {
            int index = _historyRibbon.Controls.GetChildIndex(item.Thumbnail);
            _historyRibbon.Controls.Remove(item.Thumbnail);

            if (_currentEvidence == item)
            {
                if (_historyRibbon.Controls.Count > 1)
                    SelectEvidence((EvidenceItem)((PictureBox)_historyRibbon.Controls[_historyRibbon.Controls.Count - 1]).Tag);
                else
                    SelectEvidence(null);
            }

            if (!isUndoRedo && !_isLoadingProject)
            {
                _projectUndoStack.Push(new ProjectAction
                {
                    UndoAction = () => InsertEvidenceToUI(item, index, true),
                    RedoAction = () => RemoveEvidenceFromUI(item, true)
                });
                _projectRedoStack.Clear();
            }
            ReindexHistory(); TriggerAutoSave();
        }

        private void MoveEvidenceInUI(EvidenceItem item, int oldIndex, int newIndex, bool isUndoRedo)
        {
            _historyRibbon.Controls.SetChildIndex(item.Thumbnail, newIndex);

            if (!isUndoRedo && !_isLoadingProject)
            {
                _projectUndoStack.Push(new ProjectAction
                {
                    UndoAction = () => MoveEvidenceInUI(item, newIndex, oldIndex, true),
                    RedoAction = () => MoveEvidenceInUI(item, oldIndex, newIndex, true)
                });
                _projectRedoStack.Clear();
            }
            ReindexHistory(); TriggerAutoSave();
        }

        private void Thumbnail_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is PictureBox pb && pb.Tag is EvidenceItem item)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                int index = _historyRibbon.Controls.GetChildIndex(pb);

                int visualIndex = index > 0 ? index : 1;

                e.Graphics.FillEllipse(new SolidBrush(Color.FromArgb(200, 0, 120, 215)), 4, 4, 22, 22);
                e.Graphics.DrawString(visualIndex.ToString(), new Font("Segoe UI", 9, FontStyle.Bold), Brushes.White, new Point(8, 7));
                e.Graphics.FillEllipse(new SolidBrush(Color.FromArgb(220, 220, 53, 69)), new Rectangle(pb.Width - 28, 4, 22, 22));
                e.Graphics.DrawString("X", new Font("Segoe UI", 8, FontStyle.Bold), Brushes.White, new Point(pb.Width - 23, 8));

                if (item.IsEvidenceOnly)
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(220, 255, 140, 0)), new Rectangle(2, pb.Height - 18, 95, 16));
                    e.Graphics.DrawString($"👁️ {LanguageManager.GetString("EvidenceOnly")}", new Font("Segoe UI", 7, FontStyle.Bold), Brushes.White, new Point(4, pb.Height - 17));
                }
            }
        }

        private void Thumbnail_MouseDown(object? sender, MouseEventArgs e)
        {
            if (sender is PictureBox pb && pb.Tag is EvidenceItem item)
            {
                if (e.X >= pb.Width - 30 && e.Y <= 30)
                {
                    RemoveEvidenceFromUI(item, false);
                    return;
                }

                if (_stepNoteTextBox.Focused) TriggerAutoSave();

                SelectEvidence(item);
                if (e.Button == MouseButtons.Left && e.Clicks == 1) pb.DoDragDrop(pb, DragDropEffects.Move);
            }
        }

        private void HistoryRibbon_DragDrop(object sender, DragEventArgs e)
        {
            var draggedThumb = (PictureBox)e.Data.GetData(typeof(PictureBox));
            if (draggedThumb == null) return;
            Point p = _historyRibbon.PointToClient(new Point(e.X, e.Y));
            var target = _historyRibbon.GetChildAtPoint(p);

            if (target != null && target != draggedThumb && target != _templateThumbnail && draggedThumb != _templateThumbnail)
            {
                int oldIndex = _historyRibbon.Controls.GetChildIndex(draggedThumb);
                int targetIndex = _historyRibbon.Controls.GetChildIndex(target);

                MoveEvidenceInUI((EvidenceItem)draggedThumb.Tag, oldIndex, targetIndex, false);
            }
        }

        private void ReindexHistory() { foreach (Control c in _historyRibbon.Controls) c.Invalidate(); }

        private void TriggerAutoSave()
        {
            if (_isLoadingProject || !_isAutoSaveEnabled) return;

            _hasUnsavedChanges = true;
            if (!string.IsNullOrEmpty(_currentProjectPath))
            {
                SaveProjectInternal(_currentProjectPath);
                _hasUnsavedChanges = false;
            }
        }

        private void NewProject()
        {
            if (_hasUnsavedChanges && _historyRibbon.Controls.Count > 1)
            {
                var r = MessageBox.Show(LanguageManager.GetString("MsgSaveBeforeNew"), LanguageManager.GetString("TitleNew"), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (r == DialogResult.Cancel) return;
                if (r == DialogResult.Yes) SaveProjectCurrent();
            }

            _currentProjectData = new LiteFlowProjectData();
            _currentProjectData.ReportLayout = _defaultLayoutMode;
            _currentProjectData.MobileColumns = _defaultMobileColumns;

            _currentProjectData.TemplatePath = (!string.IsNullOrEmpty(_defaultTemplatePath) && File.Exists(_defaultTemplatePath) && MessageBox.Show(LanguageManager.GetString("MsgUseDefaultTemplate"), LanguageManager.GetString("TitleTemplate"), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) ? _defaultTemplatePath : "";

            ClearEvidenceHistory();

            _templateThumbnail.Invalidate();
            _currentProjectPath = "";
            _hasUnsavedChanges = false;

            _isAutoSaveEnabled = false;
            UpdateAutoSaveUI();
            UpdateProjectNameUI();
        }

        private void SaveProjectCurrent()
        {
            if (string.IsNullOrEmpty(_currentProjectPath))
            {
                SaveProjectAs();
            }
            else
            {
                SaveProjectInternal(_currentProjectPath);
                _hasUnsavedChanges = false;

                _isAutoSaveEnabled = true;
                UpdateAutoSaveUI();
                UpdateProjectNameUI();
                MessageBox.Show(LanguageManager.GetString("MsgProjectSaved"), LanguageManager.GetString("TitleLiteFlow"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SaveProjectAs()
        {
            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Projeto LiteFlow (*.lflow)|*.lflow" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    _currentProjectPath = sfd.FileName;

                    if (string.IsNullOrWhiteSpace(_currentProjectData.FileName))
                    {
                        _currentProjectData.FileName = Path.GetFileNameWithoutExtension(_currentProjectPath);
                    }

                    SaveProjectInternal(_currentProjectPath);
                    _hasUnsavedChanges = false;

                    _isAutoSaveEnabled = true;
                    UpdateAutoSaveUI();
                    UpdateProjectNameUI();

                    MessageBox.Show(LanguageManager.GetString("MsgProjectSavedShort"), LanguageManager.GetString("TitleLiteFlow"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void OpenProject()
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Projeto (*.lflow)|*.lflow" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var data = ProjectService.LoadProject(ofd.FileName);
                        if (data != null)
                        {
                            ClearEvidenceHistory();
                            _currentProjectPath = ofd.FileName;
                            _currentProjectData = data;

                            if (string.IsNullOrWhiteSpace(_currentProjectData.FileName))
                            {
                                _currentProjectData.FileName = Path.GetFileNameWithoutExtension(_currentProjectPath);
                            }

                            _templateThumbnail.Invalidate();

                            _isLoadingProject = true;
                            foreach (var step in data.Steps)
                            {
                                string path = Path.Combine(_sessionTempDir, $"img_{Guid.NewGuid():N}.png");
                                File.WriteAllBytes(path, Convert.FromBase64String(step.ImageDataBase64));
                                step.ImageDataBase64 = null!;

                                AddToHistoryFromDisk(path, null, step.Note, step.TextBelowImage, step.IsEvidenceOnly);
                            }

                            data.Steps.Clear();
                            GC.Collect(2, GCCollectionMode.Forced, true);

                            _isLoadingProject = false;
                            _hasUnsavedChanges = false;

                            _isAutoSaveEnabled = true;
                            UpdateAutoSaveUI();
                            UpdateProjectNameUI();
                        }
                    }
                    catch (Exception ex)
                    {
                        _isLoadingProject = false;
                        MessageBox.Show(string.Format(LanguageManager.GetString("MsgError"), ex.Message), LanguageManager.GetString("TitleError"));
                    }
                }
            }
        }

        private List<EvidenceItem> GetItems()
        {
            var list = new List<EvidenceItem>();
            foreach (PictureBox pb in _historyRibbon.Controls)
            {
                if (pb.Tag is EvidenceItem item) list.Add(item);
            }
            return list;
        }

        private bool PasteImageFromClipboard()
        {
            if (Clipboard.ContainsImage())
            {
                using (Image img = Clipboard.GetImage())
                {
                    if (img != null)
                    {
                        Bitmap bmp = new Bitmap(img);
                        if (_currentEvidence != null)
                        {
                            _editorCore?.LoadImage(bmp);
                            _currentEvidence.Image?.Dispose();
                            _currentEvidence.Image = bmp;

                            _currentEvidence.Thumbnail.Image?.Dispose();
                            _currentEvidence.Thumbnail.Image = CreateThumbnail(bmp);
                            _currentEvidence.Thumbnail.Invalidate();

                            if (!string.IsNullOrEmpty(_currentEvidence.DiskPath))
                            {
                                var cloneToSave = new Bitmap(bmp);
                                string path = _currentEvidence.DiskPath;
                                Task.Run(() => {
                                    try { cloneToSave.Save(path, ImageFormat.Png); } catch { } finally { cloneToSave.Dispose(); }
                                });
                            }
                            TriggerAutoSave();
                        }
                        else
                        {
                            Task.Run(() => {
                                string path = Path.Combine(_sessionTempDir, $"img_{Guid.NewGuid():N}.png");
                                bmp.Save(path, ImageFormat.Png);
                                this.BeginInvoke(new Action(() => { AddToHistoryFromDisk(path, bmp); }));
                            });
                        }
                        return true;
                    }
                }
            }
            return false;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.V))
            {
                if (!_stepNoteTextBox.Focused && !_floatingTextBox.Focused)
                {
                    if (PasteImageFromClipboard()) return true;
                }
            }

            if (keyData == (Keys.Control | Keys.S)) { SaveProjectCurrent(); return true; }
            if (keyData == (Keys.Control | Keys.Z)) { PerformUndo(); return true; }
            if (keyData == (Keys.Control | Keys.Y)) { PerformRedo(); return true; }

            if (keyData == (Keys.Control | Keys.Oemplus) || keyData == (Keys.Control | Keys.Add)) { _trkThickness.Value = Math.Min(20, _trkThickness.Value + 1); if (_editorCore != null) _editorCore.CurrentThickness = _trkThickness.Value; SaveSettings(); return true; }
            if (keyData == (Keys.Control | Keys.OemMinus) || keyData == (Keys.Control | Keys.Subtract)) { _trkThickness.Value = Math.Max(1, _trkThickness.Value - 1); if (_editorCore != null) _editorCore.CurrentThickness = _trkThickness.Value; SaveSettings(); return true; }

            if (keyData == Keys.Enter) { _editorCore?.ConfirmCrop(); return true; }
            if (keyData == Keys.Escape) { _editorCore?.CancelCurrentAction(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void BtnImportWord_Click(object? sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Word (*.docx)|*.docx" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    _currentProjectData.TemplatePath = ofd.FileName;
                    if (MessageBox.Show(LanguageManager.GetString("MsgSetAsDefault"), LanguageManager.GetString("TitleDefault"), MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Directory.CreateDirectory(_templatesDir);
                        string destPath = Path.Combine(_templatesDir, Path.GetFileName(ofd.FileName));
                        File.Copy(ofd.FileName, destPath, true);
                        _defaultTemplatePath = destPath;
                        SaveSettings();
                        MessageBox.Show(LanguageManager.GetString("MsgTemplateSaved"));
                    }
                    TriggerAutoSave();
                    _templateThumbnail.Invalidate();
                }
            }
        }

        private void BtnExportWord_Click(object? sender, EventArgs e)
        {
            var itemsToExport = GetItems();
            if (itemsToExport.Count == 0) return;

            string templateToUse = !string.IsNullOrEmpty(_currentProjectData.TemplatePath) ? _currentProjectData.TemplatePath : _defaultTemplatePath;

            if (string.IsNullOrEmpty(templateToUse) || !File.Exists(templateToUse))
                MessageBox.Show(LanguageManager.GetString("MsgExportNoTemplate"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (TemplateDataForm form = new TemplateDataForm(_defaultQAName, _defaultPrefix, _currentProjectData))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    form.UpdateProjectData(_currentProjectData);
                    if (form.SaveQADefault) _defaultQAName = _currentProjectData.QAName;
                    if (form.SavePrefixDefault) _defaultPrefix = _currentProjectData.FilePrefix;

                    _defaultLayoutMode = _currentProjectData.ReportLayout;
                    _defaultMobileColumns = _currentProjectData.MobileColumns;

                    SaveSettings();
                    TriggerAutoSave();

                    string safeFileName = string.IsNullOrWhiteSpace(form.SuggestedFileName) ? "Evidencias_LiteFlow" : form.SuggestedFileName;

                    using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Word (*.docx)|*.docx", FileName = safeFileName + ".docx" })
                    {
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            try
                            {
                                this.Cursor = Cursors.WaitCursor;

                                ExportService.ExportToWord(_currentProjectData, itemsToExport, sfd.FileName, form.GetTags());

                                this.Cursor = Cursors.Default;
                                MessageBox.Show(string.Format(LanguageManager.GetString("MsgExportSuccessWord"), sfd.FileName), LanguageManager.GetString("TitleSuccess"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex) { this.Cursor = Cursors.Default; MessageBox.Show(string.Format(LanguageManager.GetString("MsgError"), ex.Message), LanguageManager.GetString("TitleError"), MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        }
                    }
                }
            }
        }

        private void BtnApplyToPdf_Click(object? sender, EventArgs e)
        {
            var itemsToExport = GetItems();
            if (itemsToExport.Count == 0) return;

            string templateToUse = !string.IsNullOrEmpty(_currentProjectData.TemplatePath) ? _currentProjectData.TemplatePath : _defaultTemplatePath;
            if (string.IsNullOrEmpty(templateToUse) || !File.Exists(templateToUse))
                MessageBox.Show(LanguageManager.GetString("MsgExportNoTemplate"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Information);

            using (TemplateDataForm form = new TemplateDataForm(_defaultQAName, _defaultPrefix, _currentProjectData))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    form.UpdateProjectData(_currentProjectData);
                    if (form.SaveQADefault) _defaultQAName = _currentProjectData.QAName;
                    if (form.SavePrefixDefault) _defaultPrefix = _currentProjectData.FilePrefix;

                    _defaultLayoutMode = _currentProjectData.ReportLayout;
                    _defaultMobileColumns = _currentProjectData.MobileColumns;

                    SaveSettings();
                    TriggerAutoSave();

                    string safeFileName = string.IsNullOrWhiteSpace(form.SuggestedFileName) ? "Evidencias_LiteFlow" : form.SuggestedFileName;

                    using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = safeFileName + ".pdf" })
                    {
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            try
                            {
                                this.Cursor = Cursors.WaitCursor;

                                ExportService.ExportToPdf(_currentProjectData, itemsToExport, sfd.FileName, form.GetTags());

                                this.Cursor = Cursors.Default;
                                MessageBox.Show(LanguageManager.GetString("MsgExportSuccessPdf"), LanguageManager.GetString("TitleLiteFlowPdf"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex) { this.Cursor = Cursors.Default; MessageBox.Show(string.Format(LanguageManager.GetString("MsgErrorPdf"), ex.Message), LanguageManager.GetString("TitleError"), MessageBoxButtons.OK, MessageBoxIcon.Error); }
                        }
                    }
                }
            }
        }
    }
}