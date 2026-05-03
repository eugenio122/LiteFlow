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

namespace LiteFlow.UI
{
    public partial class LiteFlowUI : UserControl, ILitePlugin
    {
        public string Name => "LiteFlow";
        public string Version => "1.0.0 (Spatial Memory Management)";

        private IEventBus? _eventBus;
        private ILiteHostContext? _hostContext;

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
        private bool _isAutoSaveEnabled = false;

        private LayoutMode _defaultLayoutMode = LayoutMode.Padrao;
        private int _defaultMobileColumns = 2;

        private bool _isLoadingProject = false;
        private bool _hasUnsavedChanges = false;
        private bool _firstLoadDone = false;
        private bool _isDarkMode = false;

        private System.Windows.Forms.Timer _autoSaveTimer = null!;
        private bool _isSavingInBackground = false;

        private Dictionary<EvidenceItem, string> _evidenceDiskPaths = new Dictionary<EvidenceItem, string>();

        private ImageEditorCore _editorCore = null!;

        private class ProjectAction
        {
            public Action UndoAction { get; set; } = null!;
            public Action RedoAction { get; set; } = null!;
        }
        private Stack<ProjectAction> _projectUndoStack = new Stack<ProjectAction>();
        private Stack<ProjectAction> _projectRedoStack = new Stack<ProjectAction>();

        // Layout Base UI
        private TableLayoutPanel _mainTable = null!;
        private ToolStrip _topToolbar = null!;
        private ToolStripButton _btnToggleCapture = null!;
        private ToolStripButton _btnAutoSaveToggle = null!;
        private ToolStripLabel _lblProjectName = null!;
        private ToolStripButton _btnRestart = null!;

        private Panel _editorContainer = null!;
        private PictureBox _mainCanvas = null!;
        private Panel _propertiesPanel = null!;
        private FlowLayoutPanel _propertiesScrollPanel = null!;

        // Elementos do Histórico
        private Panel _historyWrapper = null!;
        private FlowLayoutPanel _historyHeaderPanel = null!;
        private Label _lblHistoryTitle = null!;
        private Button _btnAddBlank = null!;
        private Button _btnPaste = null!;
        private FlowLayoutPanel _historyRibbon = null!;
        private PictureBox _templateThumbnail = null!;

        private Button _btnColorPicker = null!;
        private TrackBar _trkThickness = null!;
        private TextBox _floatingTextBox = null!;
        private RichTextBox _stepNoteTextBox = null!;
        private CheckBox _chkTextBelowStep = null!;
        private CheckBox _chkEvidenceOnly = null!;
        private ToolStripComboBox _cmbFont = null!;
        private ToolStripComboBox _cmbSize = null!;

        // Propriedades UI (Barra Lateral)
        private bool _isProgrammaticUpdate = false;
        private Button _btnImportTemplate = null!;
        private TextBox _txtPropPrefix = null!;
        private CheckBox _chkPropDefaultPrefix = null!;
        private RichTextBox _txtPropFileName = null!;
        private RichTextBox _txtPropTestCase = null!;
        private TextBox _txtPropQA = null!;
        private CheckBox _chkPropDefaultQA = null!;
        private TextBox _txtPropDate = null!;
        private RichTextBox _txtPropComments = null!;
        private ComboBox _cmbPropLayout = null!;
        private NumericUpDown _numPropCols = null!;

        public LiteFlowUI()
        {
            _configPath = Path.Combine(_baseDir, "settings.ini");
            _templatesDir = Path.Combine(_baseDir, "Templates");

            _sessionTempDir = Path.Combine(Path.GetTempPath(), "LiteFlowSession_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_sessionTempDir);

            _autoSaveTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            _autoSaveTimer.Tick += AutoSaveTimer_Tick;

            LoadLanguageOnly();
            InitializeComponent();
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
                        Task.Run(() => { try { cloneToSave.Save(path, ImageFormat.Png); } catch { } finally { cloneToSave.Dispose(); } });
                    }
                    TriggerAutoSave();
                }
            };
        }

        public void Initialize(ILiteHostContext hostContext, IEventBus eventBus, string currentLanguage)
        {
            _hostContext = hostContext;
            _eventBus = eventBus;
            LanguageManager.CurrentLanguage = currentLanguage;

            _eventBus.Subscribe<ImageCapturedEvent>(OnImageReceived);
            _eventBus.Subscribe<ThemeChangedEvent>(OnThemeChanged);

            if (_btnToggleCapture != null)
            {
                _btnToggleCapture.Enabled = true;
                _btnToggleCapture.Text = _isRecording ? LanguageManager.GetString("BtnCapturing") : LanguageManager.GetString("BtnPaused");
                _btnToggleCapture.BackColor = _isRecording ? Color.LightGreen : Color.LightCoral;
                _btnToggleCapture.ForeColor = Color.Black;
                _btnToggleCapture.ToolTipText = LanguageManager.GetString("TooltipPause");

                _eventBus.Publish(new RecordingStateChangedEvent(_isRecording));
            }

            UpdateProjectNameUI();
            UpdateAutoSaveUI();
            FeedTestCaseToHostContext();
        }

        public UserControl GetSettingsUI() { return this; }

        public void Shutdown()
        {
            if (_hasUnsavedChanges && _historyRibbon.Controls.Count > 1)
            {
                var r = MessageBox.Show(LanguageManager.GetString("MsgUnsavedShutdown"), LanguageManager.GetString("Warning"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (r == DialogResult.Yes) SaveProjectCurrent();
            }

            ClearEvidenceHistory();
            try { if (Directory.Exists(_sessionTempDir)) Directory.Delete(_sessionTempDir, true); } catch { }
        }

        private void OnImageReceived(ImageCapturedEvent evt)
        {
            if (evt?.CapturedImage == null || !_isRecording || this.IsDisposed) return;
            Bitmap clonedImage = new Bitmap(evt.CapturedImage);

            if (!this.IsHandleCreated) _pendingImages.Enqueue(clonedImage);
            else
            {
                Task.Run(() => {
                    string diskPath = Path.Combine(_sessionTempDir, $"img_{Guid.NewGuid():N}.png");
                    clonedImage.Save(diskPath, ImageFormat.Png);
                    this.BeginInvoke(new Action(() => {
                        if (!this.IsDisposed) AddToHistoryFromDisk(diskPath, clonedImage, "", false, false, evt.StepId);
                        else clonedImage.Dispose();
                    }));
                });
            }
        }

        public void ReceiveImage(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath) || !_isRecording || this.IsDisposed) return;
            this.BeginInvoke(new Action(() => { if (!this.IsDisposed) AddToHistoryFromDisk(imagePath, null); }));
        }

        private void AddToHistoryFromDisk(string diskPath, Bitmap? alreadyLoadedImage = null, string note = "", bool textBelow = false, bool isEvidenceOnly = false, string stepId = "")
        {
            if (!File.Exists(diskPath)) return;

            Bitmap activeImg = alreadyLoadedImage ?? LoadImageFromDisk(diskPath);
            Bitmap thumbImg = CreateThumbnail(activeImg);

            var item = new EvidenceItem
            {
                StepId = string.IsNullOrEmpty(stepId) ? Guid.NewGuid().ToString("N") : stepId,
                Image = activeImg,
                Note = note,
                TextBelowImage = textBelow,
                IsEvidenceOnly = isEvidenceOnly,
                DiskPath = diskPath
            };

            Color thumbBg = _isDarkMode ? Color.FromArgb(60, 60, 60) : Color.White;
            var thumbnailPb = new PictureBox { Width = 110, Height = 70, SizeMode = PictureBoxSizeMode.Zoom, Image = thumbImg, BackColor = thumbBg, Padding = new Padding(2), Margin = new Padding(5), Cursor = Cursors.Hand, Tag = item };

            item.Thumbnail = thumbnailPb;
            thumbnailPb.Paint += Thumbnail_Paint;
            thumbnailPb.MouseDown += Thumbnail_MouseDown;

            InsertEvidenceToUI(item, _historyRibbon.Controls.Count, false);
            if (_stepNoteTextBox != null && _stepNoteTextBox.Focused) TriggerAutoSave();

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
            int w = 110;
            int h = (int)((110.0f / original.Width) * original.Height);
            if (h <= 0) h = 70;

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

        // =========================================================================
        // GESTÃO ESPACIAL DE MEMÓRIA (Mantém na RAM: Alvo, Esquerda e Direita)
        // =========================================================================
        private void ManageMemoryFocus(EvidenceItem activeItem)
        {
            var allItems = GetItems(); // Lê a ordem visual atual
            int activeIndex = allItems.IndexOf(activeItem);
            if (activeIndex == -1) return;

            // Define o grupo VIP (Regra de 3)
            var itemsToKeep = new HashSet<EvidenceItem>();
            itemsToKeep.Add(activeItem);
            if (activeIndex > 0) itemsToKeep.Add(allItems[activeIndex - 1]); // Vizinho da esquerda
            if (activeIndex < allItems.Count - 1) itemsToKeep.Add(allItems[activeIndex + 1]); // Vizinho da direita

            // Varre todas as imagens e liberta/carrega conforme necessário
            foreach (var item in allItems)
            {
                if (itemsToKeep.Contains(item))
                {
                    // Carrega do disco caso seja vizinho e não esteja na RAM
                    if (item.Image == null && !string.IsNullOrEmpty(item.DiskPath) && File.Exists(item.DiskPath))
                    {
                        item.Image = LoadImageFromDisk(item.DiskPath);
                    }
                }
                else
                {
                    // Fora do grupo VIP: Destrói o Bitmap da memória RAM
                    if (item.Image != null)
                    {
                        item.Image.Dispose();
                        item.Image = null!;
                    }
                }
            }

            // GC leve só para limpar pequenos ponteiros
            GC.Collect(2, GCCollectionMode.Optimized, false);
        }

        private void SelectEvidence(EvidenceItem? item)
        {
            if (_currentEvidence != null && _currentEvidence.Thumbnail != null)
            {
                _currentEvidence.Thumbnail.BackColor = _isDarkMode ? Color.FromArgb(60, 60, 60) : Color.White;
            }

            _currentEvidence = item;

            if (_currentEvidence == null)
            {
                _editorCore?.LoadImage(new Bitmap(10, 10));
                if (_stepNoteTextBox != null)
                {
                    _stepNoteTextBox.Text = ""; _stepNoteTextBox.Enabled = false;
                    _chkTextBelowStep.Checked = false; _chkTextBelowStep.Enabled = false;
                    if (_chkEvidenceOnly != null) { _chkEvidenceOnly.Checked = false; _chkEvidenceOnly.Enabled = false; }
                }
                return;
            }

            _currentEvidence.Thumbnail.BackColor = Color.FromArgb(0, 120, 215);

            // A MÁGICA DE GESTÃO ACONTECE AQUI ANTES DE MOSTRAR A IMAGEM
            ManageMemoryFocus(_currentEvidence);

            _editorCore?.LoadImage(_currentEvidence.Image); // Seguro porque o ManageMemoryFocus acabou de a carregar

            if (_stepNoteTextBox != null)
            {
                _stepNoteTextBox.Enabled = true; _stepNoteTextBox.Text = _currentEvidence.Note;
                _chkTextBelowStep.Enabled = true; _chkTextBelowStep.Checked = _currentEvidence.TextBelowImage;
                _chkEvidenceOnly.Enabled = true; _chkEvidenceOnly.Checked = _currentEvidence.IsEvidenceOnly;
            }

            foreach (ToolStripItem t in _topToolbar.Items)
                if (t is ToolStripButton b && b.Tag is EditorTool et && et == EditorTool.Select) { b.PerformClick(); break; }
        }

        private void FeedTestCaseToHostContext()
        {
            if (!string.IsNullOrEmpty(_currentProjectData.TestCaseName))
                _hostContext?.SetSessionMetadata("CurrentTestCaseName", _currentProjectData.TestCaseName);
        }

        private void PerformUndo()
        {
            if (_editorCore != null && _editorCore.CanUndo) { _editorCore.Undo(); }
            else if (_projectUndoStack.Count > 0) { var action = _projectUndoStack.Pop(); action.UndoAction(); _projectRedoStack.Push(action); }
        }

        private void PerformRedo()
        {
            if (_editorCore != null && _editorCore.CanRedo) { _editorCore.Redo(); }
            else if (_projectRedoStack.Count > 0) { var action = _projectRedoStack.Pop(); action.RedoAction(); _projectUndoStack.Push(action); }
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
                                Task.Run(() => { try { cloneToSave.Save(path, ImageFormat.Png); } catch { } finally { cloneToSave.Dispose(); } });
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
                if ((_stepNoteTextBox == null || !_stepNoteTextBox.Focused) &&
                    (_floatingTextBox == null || !_floatingTextBox.Focused) &&
                    (_txtPropComments == null || !_txtPropComments.Focused) &&
                    (_txtPropFileName == null || !_txtPropFileName.Focused) &&
                    (_txtPropTestCase == null || !_txtPropTestCase.Focused))
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
    }
}