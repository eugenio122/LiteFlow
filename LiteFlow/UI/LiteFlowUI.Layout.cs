using LiteFlow.Core;
using LiteFlow.Models;
using LiteTools.Interfaces;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LiteFlow.UI
{
    public partial class LiteFlowUI
    {
        // =========================================================
        // WIN32 API: Força o Windows a pintar as Scrollbars no modo Escuro
        // =========================================================
        [DllImport("uxtheme.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hwnd, string pszSubAppName, string? pszSubIdList);

        private void ApplyScrollbarTheme(Control ctrl, bool isDark)
        {
            if (ctrl == null) return;
            if (!ctrl.IsHandleCreated)
            {
                try { var h = ctrl.Handle; } catch { }
            }
            try
            {
                SetWindowTheme(ctrl.Handle, isDark ? "DarkMode_Explorer" : "Explorer", null);
            }
            catch { }
        }
        // =========================================================

        // =========================================================
        // CUSTOM RENDERER: Resolve o Hover azul matando os gradientes
        // =========================================================
        private class DarkModeColorTable : ProfessionalColorTable
        {
            private Color HoverColor = Color.FromArgb(65, 65, 68);
            private Color PressedColor = Color.FromArgb(85, 85, 88);
            private Color BgColor = Color.FromArgb(45, 45, 48);

            // HOVER (Rato por cima) - Mata os gradientes brancos nativos
            public override Color ButtonSelectedHighlight => HoverColor;
            public override Color ButtonSelectedBorder => HoverColor;
            public override Color ButtonSelectedGradientBegin => HoverColor;
            public override Color ButtonSelectedGradientMiddle => HoverColor;
            public override Color ButtonSelectedGradientEnd => HoverColor;

            // PRESSED (Clicar no botão)
            public override Color ButtonPressedHighlight => PressedColor;
            public override Color ButtonPressedBorder => PressedColor;
            public override Color ButtonPressedGradientBegin => PressedColor;
            public override Color ButtonPressedGradientMiddle => PressedColor;
            public override Color ButtonPressedGradientEnd => PressedColor;

            // CHECKED (Ferramenta atualmente selecionada, ex: Seta)
            public override Color ButtonCheckedHighlight => PressedColor;
            public override Color ButtonCheckedHighlightBorder => PressedColor;
            public override Color ButtonCheckedGradientBegin => PressedColor;
            public override Color ButtonCheckedGradientMiddle => PressedColor;
            public override Color ButtonCheckedGradientEnd => PressedColor;

            // FUNDO DA NAVBAR
            public override Color ToolStripBorder => BgColor;
            public override Color ToolStripGradientBegin => BgColor;
            public override Color ToolStripGradientMiddle => BgColor;
            public override Color ToolStripGradientEnd => BgColor;
            public override Color ToolStripDropDownBackground => BgColor;
        }
        // =========================================================

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(240, 240, 240);
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            this.DoubleBuffered = true;

            SetupToolbar();

            _mainTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.FromArgb(240, 240, 240)
            };

            _mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 155F));
            _mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 330F));

            SetupHistoryRibbon();
            SetupEditorArea();
            SetupPropertiesPanel();

            this.Controls.Add(_mainTable);
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

                ApplyTheme(_isDarkMode);

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

        private void OnThemeChanged(ThemeChangedEvent e)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            this.BeginInvoke(new Action(() => ApplyTheme(e.IsDarkMode)));
        }

        private void ApplyTheme(bool isDark)
        {
            _isDarkMode = isDark;
            SaveSettings();

            Color bgMain = isDark ? Color.FromArgb(30, 30, 30) : Color.FromArgb(240, 240, 240);
            Color surface = isDark ? Color.FromArgb(45, 45, 48) : Color.White;
            Color textMain = isDark ? Color.WhiteSmoke : Color.Black;
            Color inputBg = isDark ? Color.FromArgb(60, 60, 60) : Color.White;

            this.BackColor = bgMain;
            _mainTable.BackColor = bgMain;

            _topToolbar.BackColor = surface;
            _topToolbar.ForeColor = textMain;

            // Aplica a nova classe DarkModeColorTable
            if (isDark)
            {
                _topToolbar.Renderer = new ToolStripProfessionalRenderer(new DarkModeColorTable()) { RoundedEdges = false };
            }
            else
            {
                _topToolbar.Renderer = new ToolStripProfessionalRenderer() { RoundedEdges = false };
            }

            if (_cmbFont != null) { _cmbFont.BackColor = inputBg; _cmbFont.ForeColor = textMain; }
            if (_cmbSize != null) { _cmbSize.BackColor = inputBg; _cmbSize.ForeColor = textMain; }
            if (_lblProjectName != null) { _lblProjectName.ForeColor = textMain; }
            if (_btnToggleCapture != null) { _btnToggleCapture.ForeColor = Color.Black; }

            if (_historyWrapper != null) _historyWrapper.BackColor = surface;
            if (_historyHeaderPanel != null) _historyHeaderPanel.BackColor = surface;
            if (_historyRibbon != null) _historyRibbon.BackColor = surface;
            if (_lblHistoryTitle != null) _lblHistoryTitle.ForeColor = textMain;

            if (_btnAddBlank != null) { _btnAddBlank.BackColor = inputBg; _btnAddBlank.ForeColor = textMain; }
            if (_btnPaste != null) { _btnPaste.BackColor = inputBg; _btnPaste.ForeColor = textMain; }

            if (_templateThumbnail != null)
            {
                _templateThumbnail.BackColor = isDark ? inputBg : Color.WhiteSmoke;
                _templateThumbnail.Invalidate();
            }

            if (_historyRibbon != null)
            {
                foreach (Control c in _historyRibbon.Controls)
                {
                    if (c is PictureBox pb)
                    {
                        if (pb.Tag is EvidenceItem item && item != _currentEvidence)
                            pb.BackColor = isDark ? inputBg : Color.White;
                        pb.Invalidate();
                    }
                }
            }

            _propertiesPanel.BackColor = surface;

            if (_chkEvidenceOnly != null) _chkEvidenceOnly.ForeColor = textMain;
            if (_chkTextBelowStep != null) _chkTextBelowStep.ForeColor = textMain;
            if (_chkPropDefaultQA != null) _chkPropDefaultQA.ForeColor = textMain;
            if (_chkPropDefaultPrefix != null) _chkPropDefaultPrefix.ForeColor = textMain;

            if (_stepNoteTextBox != null) { _stepNoteTextBox.BackColor = inputBg; _stepNoteTextBox.ForeColor = textMain; }

            foreach (Control ctrl in _propertiesPanel.Controls)
            {
                if (ctrl is FlowLayoutPanel flp)
                {
                    flp.BackColor = surface;
                    foreach (Control inner in flp.Controls)
                    {
                        if (inner is Label lbl) lbl.ForeColor = textMain;
                        else if (inner is TextBox txt) { txt.BackColor = inputBg; txt.ForeColor = textMain; }
                        else if (inner is RichTextBox rtb) { rtb.BackColor = inputBg; rtb.ForeColor = textMain; }
                        else if (inner is ComboBox cmb) { cmb.BackColor = inputBg; cmb.ForeColor = textMain; }
                        else if (inner is NumericUpDown num) { num.BackColor = inputBg; num.ForeColor = textMain; }
                        else if (inner is CheckBox chk) { chk.ForeColor = textMain; }
                    }
                }
            }

            if (_editorContainer.Controls.Count > 0 && _editorContainer.Controls[0] is Panel imgContainer)
            {
                imgContainer.BackColor = bgMain;
            }

            ApplyScrollbarTheme(_historyRibbon!, isDark);
            ApplyScrollbarTheme(_propertiesScrollPanel!, isDark);
            ApplyScrollbarTheme(_stepNoteTextBox!, isDark);
            ApplyScrollbarTheme(_txtPropFileName!, isDark);
            ApplyScrollbarTheme(_txtPropTestCase!, isDark);
            ApplyScrollbarTheme(_txtPropComments!, isDark);
        }

        private void LoadLanguageOnly()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    var lines = File.ReadAllLines(_configPath);
                    if (lines.Length > 12 && !string.IsNullOrEmpty(lines[12]))
                        LanguageManager.CurrentLanguage = lines[12];
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
                    if (lines.Length > 1) { _btnColorPicker.BackColor = ColorTranslator.FromHtml(lines[1]); }
                    if (lines.Length > 2 && int.TryParse(lines[2], out int thick)) { _trkThickness.Value = Math.Max(1, Math.Min(20, thick)); }
                    if (lines.Length > 3) _cmbFont.Text = lines[3];
                    if (lines.Length > 4) _cmbSize.Text = lines[4];
                    if (lines.Length > 5) { _defaultTemplatePath = lines[5]; _currentProjectData.TemplatePath = lines[5]; }
                    if (lines.Length > 6) { _defaultQAName = lines[6]; _currentProjectData.QAName = lines[6]; }
                    if (lines.Length > 7) { _defaultPrefix = lines[7]; _currentProjectData.FilePrefix = lines[7]; }
                    if (lines.Length > 9 && bool.TryParse(lines[9], out bool isDark)) { _isDarkMode = isDark; }
                    if (lines.Length > 10 && Enum.TryParse(lines[10], out LayoutMode lMode)) { _defaultLayoutMode = lMode; _currentProjectData.ReportLayout = lMode; }
                    if (lines.Length > 11 && int.TryParse(lines[11], out int cols)) { _defaultMobileColumns = cols; _currentProjectData.MobileColumns = cols; }
                }
                UpdatePropertiesPanelFromData();
            }
            catch { }
        }

        private void SaveSettings()
        {
            try
            {
                Directory.CreateDirectory(_baseDir);
                if (_chkPropDefaultQA != null && _chkPropDefaultQA.Checked) _defaultQAName = _currentProjectData.QAName;
                if (_chkPropDefaultPrefix != null && _chkPropDefaultPrefix.Checked) _defaultPrefix = _currentProjectData.FilePrefix;

                _defaultLayoutMode = _currentProjectData.ReportLayout;
                _defaultMobileColumns = _currentProjectData.MobileColumns;

                var lines = new string[] {
                    "160", ColorTranslator.ToHtml(_btnColorPicker.BackColor), _trkThickness.Value.ToString(),
                    _cmbFont.Text, _cmbSize.Text, _defaultTemplatePath, _defaultQAName, _defaultPrefix,
                    "false", _isDarkMode.ToString(), _defaultLayoutMode.ToString(), _defaultMobileColumns.ToString(), LanguageManager.CurrentLanguage
                };
                File.WriteAllLines(_configPath, lines);
                _templateThumbnail?.Invalidate();
            }
            catch { }
        }

        private void SetupToolbar()
        {
            _topToolbar = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, BackColor = Color.White, RenderMode = ToolStripRenderMode.Professional, Padding = new Padding(5) };

            _btnToggleCapture = new ToolStripButton(LanguageManager.GetString("BtnCapturing")) { BackColor = Color.LightGray, Enabled = false, ToolTipText = LanguageManager.GetString("TooltipStandalone") };
            _btnToggleCapture.Click += (s, e) => {
                _isRecording = !_isRecording;
                _btnToggleCapture.Text = _isRecording ? LanguageManager.GetString("BtnCapturing") : LanguageManager.GetString("BtnPaused");
                _btnToggleCapture.BackColor = _isRecording ? Color.LightGreen : Color.LightCoral;
                _btnToggleCapture.ForeColor = Color.Black;
                _eventBus?.Publish(new RecordingStateChangedEvent(_isRecording));
            };

            var btnNew = new ToolStripButton(LanguageManager.GetString("BtnNew")) { ToolTipText = LanguageManager.GetString("TooltipNew") }; btnNew.Click += (s, e) => NewProject();
            var btnOpen = new ToolStripButton(LanguageManager.GetString("BtnOpen")) { ToolTipText = LanguageManager.GetString("TooltipOpen") }; btnOpen.Click += (s, e) => OpenProject();
            var btnSave = new ToolStripButton(LanguageManager.GetString("BtnSave")) { ToolTipText = LanguageManager.GetString("TooltipSave") }; btnSave.Click += (s, e) => SaveProjectCurrent();
            var btnSaveAs = new ToolStripButton(LanguageManager.GetString("BtnSaveAs")) { ToolTipText = LanguageManager.GetString("TooltipSaveAs") }; btnSaveAs.Click += (s, e) => SaveProjectAs();

            _btnAutoSaveToggle = new ToolStripButton(LanguageManager.GetString("AutoSaveOff")) { ForeColor = Color.Gray, Font = new Font("Segoe UI", 9F, FontStyle.Bold), ToolTipText = LanguageManager.GetString("TooltipAutoSave") };
            _btnAutoSaveToggle.Click += (s, e) => {
                if (string.IsNullOrEmpty(_currentProjectPath)) { MessageBox.Show(LanguageManager.GetString("MsgSaveFirst"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                _isAutoSaveEnabled = !_isAutoSaveEnabled;
                UpdateAutoSaveUI();
                if (_isAutoSaveEnabled) TriggerAutoSave();
            };

            var btnUndo = new ToolStripButton("↩️") { ToolTipText = "Ctrl+Z" }; btnUndo.Click += (s, e) => PerformUndo();
            var btnRedo = new ToolStripButton("↪️") { ToolTipText = "Ctrl+Y" }; btnRedo.Click += (s, e) => PerformRedo();

            _cmbFont = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120, FlatStyle = FlatStyle.Flat };
            _cmbFont.Items.AddRange(new object[] { "Segoe UI", "Arial", "Consolas", "Times New Roman", "Verdana" });
            _cmbFont.SelectedItem = "Segoe UI";
            _cmbFont.SelectedIndexChanged += (s, e) => { if (_editorCore != null) { _editorCore.CurrentFontFamily = _cmbFont.Text; SaveSettings(); } };

            _cmbSize = new ToolStripComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 50, FlatStyle = FlatStyle.Flat };
            _cmbSize.Items.AddRange(new object[] { "10", "12", "14", "16", "18", "24", "32", "48", "72", "96" });
            _cmbSize.SelectedItem = "14";
            _cmbSize.SelectedIndexChanged += (s, e) => { if (_editorCore != null && int.TryParse(_cmbSize.Text, out int size)) { _editorCore.CurrentFontSize = size; SaveSettings(); } };

            var btnGlobalSettings = new ToolStripButton("⚙️") { ToolTipText = LanguageManager.GetString("SettingsTitle"), Alignment = ToolStripItemAlignment.Right };
            btnGlobalSettings.Click += BtnGlobalSettings_Click;

            _btnRestart = new ToolStripButton(LanguageManager.GetString("BtnRestart")) { ToolTipText = LanguageManager.GetString("TooltipRestart"), Alignment = ToolStripItemAlignment.Right };
            _btnRestart.Click += (s, e) => RestartProject();

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
                btnGlobalSettings, new ToolStripSeparator { Alignment = ToolStripItemAlignment.Right },
                _btnRestart, _lblProjectName
            };

            foreach (var item in tools)
            {
                if (item is ToolStripButton btn && btn.Tag is EditorTool)
                {
                    btn.Click += (s, e) => {
                        foreach (ToolStripItem t in _topToolbar.Items) if (t is ToolStripButton b && b.Tag is EditorTool) b.Checked = false;
                        btn.Checked = true;
                        _editorCore.CurrentTool = (EditorTool)btn.Tag;

                        _mainCanvas.Focus();

                        _editorCore.CancelCurrentAction();
                    };
                }
            }
            _topToolbar.Items.AddRange(tools);
        }

        private void UpdateAutoSaveUI()
        {
            if (string.IsNullOrEmpty(_currentProjectPath)) { _isAutoSaveEnabled = false; _btnAutoSaveToggle.Text = LanguageManager.GetString("AutoSaveOff"); _btnAutoSaveToggle.ForeColor = Color.Gray; }
            else { _btnAutoSaveToggle.Text = _isAutoSaveEnabled ? LanguageManager.GetString("AutoSaveOn") : LanguageManager.GetString("AutoSaveOff"); _btnAutoSaveToggle.ForeColor = _isAutoSaveEnabled ? Color.Green : Color.Orange; }
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
                frm.Size = new Size(420, 290);
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.FormBorderStyle = FormBorderStyle.FixedDialog;
                frm.MaximizeBox = false;
                frm.MinimizeBox = false;
                frm.BackColor = _isDarkMode ? Color.FromArgb(45, 45, 48) : Color.White;
                frm.Font = new Font("Segoe UI", 9F);

                var flpMain = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(20), WrapContents = false };

                var flpHeader = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 0, 0, 15) };
                var pbIcon = new PictureBox { Image = SystemIcons.Information.ToBitmap(), Size = new Size(32, 32), SizeMode = PictureBoxSizeMode.Zoom, Margin = new Padding(0, 0, 10, 0) };
                var flpTitle = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown };

                var lblTitle = new Label { Text = LanguageManager.GetString("SettingsHeader"), Font = new Font("Segoe UI", 12F, FontStyle.Bold), AutoSize = true, ForeColor = Color.FromArgb(0, 120, 215) };
                var lblVersion = new Label { Text = $"{LanguageManager.GetString("SettingsVersion")}{this.Version}", AutoSize = true, ForeColor = _isDarkMode ? Color.Silver : Color.DimGray };

                flpTitle.Controls.Add(lblTitle);
                flpTitle.Controls.Add(lblVersion);
                flpHeader.Controls.Add(pbIcon);
                flpHeader.Controls.Add(flpTitle);
                flpMain.Controls.Add(flpHeader);

                var lblLang = new Label { Text = LanguageManager.GetString("SettingsLanguage"), AutoSize = true, ForeColor = _isDarkMode ? Color.WhiteSmoke : Color.Black };
                var cmbLang = new ComboBox { Width = 150, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = _isDarkMode ? Color.FromArgb(60, 60, 60) : Color.White, ForeColor = _isDarkMode ? Color.WhiteSmoke : Color.Black, FlatStyle = FlatStyle.Flat };
                cmbLang.Items.AddRange(new string[] { "pt-BR", "en-US", "es-ES", "fr-FR", "de-DE", "it-IT" });

                if (cmbLang.Items.Contains(LanguageManager.CurrentLanguage)) cmbLang.SelectedItem = LanguageManager.CurrentLanguage;
                else cmbLang.SelectedIndex = 0;

                bool isPluginMode = _hostContext != null;
                if (isPluginMode)
                {
                    cmbLang.Enabled = false;
                    var lblInfo = new Label { Text = LanguageManager.GetString("LangManagedByHost"), AutoSize = true, ForeColor = Color.Gray };
                    flpMain.Controls.Add(lblLang);

                    var flpComboRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
                    flpComboRow.Controls.Add(cmbLang);
                    flpComboRow.Controls.Add(lblInfo);
                    flpMain.Controls.Add(flpComboRow);
                }
                else
                {
                    flpMain.Controls.Add(lblLang);
                    flpMain.Controls.Add(cmbLang);
                }

                var lnkGithub = new LinkLabel { Text = "GitHub: github.com/eugenio122/LiteFlow", AutoSize = true, LinkColor = Color.FromArgb(0, 120, 215), Margin = new Padding(0, 15, 0, 15) };
                lnkGithub.LinkClicked += (s, ev) => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://github.com/eugenio122/LiteFlow") { UseShellExecute = true });
                flpMain.Controls.Add(lnkGithub);

                var btnOk = new Button { Text = LanguageManager.GetString("BtnSaveClose"), Width = 120, Height = 30, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
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

                var flpFooter = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 40, Padding = new Padding(0, 0, 20, 0) };
                flpFooter.Controls.Add(btnOk);

                frm.Controls.Add(flpMain);
                frm.Controls.Add(flpFooter);
                frm.ShowDialog();
            }
        }

        private void SetupEditorArea()
        {
            _editorContainer = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(220, 220, 220) };
            _mainCanvas = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(200, 200, 200), Cursor = Cursors.Cross };
            _floatingTextBox = new TextBox { Visible = false, Multiline = true, BorderStyle = BorderStyle.FixedSingle };
            _mainCanvas.Controls.Add(_floatingTextBox);

            _mainCanvas.MouseDown += (s, e) => { if (_stepNoteTextBox != null && _stepNoteTextBox.Focused) _editorContainer.Focus(); };

            _editorContainer.Controls.Add(_mainCanvas);
            _mainTable.Controls.Add(_editorContainer, 1, 0);
        }

        private void SetupPropertiesPanel()
        {
            _propertiesPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

            var pnlExport = new Panel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(15, 10, 15, 15) };

            var btnPdf = new Button { Text = LanguageManager.GetString("BtnExportPdf"), Dock = DockStyle.Bottom, AutoSize = true, MinimumSize = new Size(0, 38), BackColor = Color.FromArgb(220, 53, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnPdf.FlatAppearance.BorderSize = 0; btnPdf.Click += BtnApplyToPdf_Click;

            var lblSpacer = new Label { Dock = DockStyle.Bottom, Height = 10 };

            var btnWordExport = new Button { Text = LanguageManager.GetString("BtnExportWord"), Dock = DockStyle.Bottom, AutoSize = true, MinimumSize = new Size(0, 38), BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand };
            btnWordExport.FlatAppearance.BorderSize = 0; btnWordExport.Click += BtnExportWord_Click;

            pnlExport.Controls.Add(btnWordExport);
            pnlExport.Controls.Add(lblSpacer);
            pnlExport.Controls.Add(btnPdf);

            _propertiesScrollPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(15, 10, 10, 10) };

            int ctrlWidth = 280;

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblProperties"), Font = new Font("Segoe UI", 10F, FontStyle.Bold), AutoSize = true, ForeColor = Color.FromArgb(64, 64, 64), Margin = new Padding(0, 0, 0, 10) });

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblColor"), AutoSize = true, ForeColor = Color.DimGray });
            _btnColorPicker = new Button { Width = ctrlWidth, Height = 30, BackColor = Color.FromArgb(0, 178, 89), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0, 0, 0, 10) };
            _btnColorPicker.Click += (s, e) => {
                using (ColorDialog cd = new ColorDialog { Color = _btnColorPicker.BackColor })
                {
                    if (cd.ShowDialog() == DialogResult.OK) { _btnColorPicker.BackColor = cd.Color; if (_editorCore != null) { _editorCore.CurrentColor = cd.Color; _floatingTextBox.ForeColor = cd.Color; SaveSettings(); } }
                }
            };
            _propertiesScrollPanel.Controls.Add(_btnColorPicker);

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblThickness"), AutoSize = true, ForeColor = Color.DimGray });
            _trkThickness = new TrackBar { Width = ctrlWidth, Minimum = 1, Maximum = 20, Value = 4, TickStyle = TickStyle.None, Height = 30, Margin = new Padding(0, 0, 0, 10) };
            _trkThickness.Scroll += (s, e) => { if (_editorCore != null) { _editorCore.CurrentThickness = _trkThickness.Value; SaveSettings(); } };
            _propertiesScrollPanel.Controls.Add(_trkThickness);

            _btnImportTemplate = new Button { Text = LanguageManager.GetString("BtnImportTemplate"), Width = ctrlWidth, AutoSize = true, MinimumSize = new Size(0, 32), BackColor = Color.SteelBlue, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0, 0, 0, 15) };
            _btnImportTemplate.FlatAppearance.BorderSize = 0;
            _btnImportTemplate.Click += BtnImportWord_Click;
            _propertiesScrollPanel.Controls.Add(_btnImportTemplate);

            _propertiesScrollPanel.Controls.Add(new Label { Width = ctrlWidth, Height = 2, BorderStyle = BorderStyle.Fixed3D, Margin = new Padding(0, 5, 0, 15) });

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("EvidenceNotes"), Font = new Font("Segoe UI", 9F, FontStyle.Bold), AutoSize = true, ForeColor = Color.FromArgb(0, 120, 215), Margin = new Padding(0, 0, 0, 5) });

            _stepNoteTextBox = new RichTextBox { Width = ctrlWidth, Height = 75, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 5) };
            var ctxMenu = new ContextMenuStrip();
            ctxMenu.Items.Add(LanguageManager.GetString("CtxPasteFormat"), null, (s, e) => _stepNoteTextBox.Paste());
            ctxMenu.Items.Add(LanguageManager.GetString("CtxPasteNoFormat"), null, (s, e) => { if (Clipboard.ContainsText()) _stepNoteTextBox.SelectedText = Clipboard.GetText(TextDataFormat.Text); });
            _stepNoteTextBox.ContextMenuStrip = ctxMenu;

            _stepNoteTextBox.TextChanged += (s, e) => { if (_currentEvidence != null) { _currentEvidence.Note = _stepNoteTextBox.Text; _hasUnsavedChanges = true; } };
            _stepNoteTextBox.Leave += (s, e) => {
                TriggerAutoSave();
                if (_currentEvidence != null) _eventBus?.Publish(new StepMetadataChangedEvent(_currentEvidence.StepId, _currentEvidence.IsEvidenceOnly, _currentEvidence.Note));
            };
            _propertiesScrollPanel.Controls.Add(_stepNoteTextBox);

            var pnlEvidChecks = new FlowLayoutPanel { Width = ctrlWidth, AutoSize = true, FlowDirection = FlowDirection.TopDown, Margin = new Padding(0, 0, 0, 10) };
            _chkEvidenceOnly = new CheckBox { Text = $"👁️ {LanguageManager.GetString("EvidenceOnly")}", AutoSize = true, ForeColor = Color.DimGray, Cursor = Cursors.Hand };
            _chkTextBelowStep = new CheckBox { Text = LanguageManager.GetString("TextBelowImage"), AutoSize = true, ForeColor = Color.DimGray, Cursor = Cursors.Hand };

            _chkTextBelowStep.CheckedChanged += (s, e) => {
                if (_currentEvidence != null && _currentEvidence.TextBelowImage != _chkTextBelowStep.Checked) { _currentEvidence.TextBelowImage = _chkTextBelowStep.Checked; _hasUnsavedChanges = true; TriggerAutoSave(); }
            };

            _chkEvidenceOnly.CheckedChanged += (s, e) => {
                if (_currentEvidence != null && _currentEvidence.IsEvidenceOnly != _chkEvidenceOnly.Checked)
                {
                    _currentEvidence.IsEvidenceOnly = _chkEvidenceOnly.Checked;
                    _currentEvidence.Thumbnail.Invalidate();
                    _hasUnsavedChanges = true; TriggerAutoSave();
                    _eventBus?.Publish(new StepMetadataChangedEvent(_currentEvidence.StepId, _currentEvidence.IsEvidenceOnly, _currentEvidence.Note));
                    ReindexHistory();
                }
            };

            pnlEvidChecks.Controls.Add(_chkEvidenceOnly);
            pnlEvidChecks.Controls.Add(_chkTextBelowStep);
            _propertiesScrollPanel.Controls.Add(pnlEvidChecks);

            _propertiesScrollPanel.Controls.Add(new Label { Width = ctrlWidth, Height = 2, BorderStyle = BorderStyle.Fixed3D, Margin = new Padding(0, 5, 0, 15) });

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblReportData"), Font = new Font("Segoe UI", 9F, FontStyle.Bold), AutoSize = true, ForeColor = Color.FromArgb(0, 120, 215), Margin = new Padding(0, 0, 0, 10) });

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblPrefix"), AutoSize = true, ForeColor = Color.DimGray });
            _txtPropPrefix = new TextBox { Width = ctrlWidth, Margin = new Padding(0, 0, 0, 2) };
            _txtPropPrefix.TextChanged += (s, e) => { if (_isProgrammaticUpdate) return; _currentProjectData.FilePrefix = _txtPropPrefix.Text; _hasUnsavedChanges = true; TriggerAutoSave(); };
            _propertiesScrollPanel.Controls.Add(_txtPropPrefix);

            _chkPropDefaultPrefix = new CheckBox { Text = "Salvar como Padrão", AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8F), Margin = new Padding(0, 0, 0, 10), Checked = true };
            _chkPropDefaultPrefix.CheckedChanged += (s, e) => { SaveSettings(); };
            _propertiesScrollPanel.Controls.Add(_chkPropDefaultPrefix);

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblFileName"), AutoSize = true, ForeColor = Color.DimGray });
            _txtPropFileName = new RichTextBox { Width = ctrlWidth, Height = 35, Multiline = true, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 10) };
            _txtPropFileName.TextChanged += (s, e) => { if (_isProgrammaticUpdate) return; _currentProjectData.FileName = _txtPropFileName.Text; _hasUnsavedChanges = true; TriggerAutoSave(); };
            var ctxFileName = new ContextMenuStrip();
            ctxFileName.Items.Add(LanguageManager.GetString("CtxPasteNoFormat"), null, (s, e) => { if (Clipboard.ContainsText()) _txtPropFileName.SelectedText = Clipboard.GetText(TextDataFormat.Text); });
            _txtPropFileName.ContextMenuStrip = ctxFileName;
            _propertiesScrollPanel.Controls.Add(_txtPropFileName);

            string lblCasoText = LanguageManager.GetString("LblTestCase");
            if (!lblCasoText.Contains("{CASO}")) lblCasoText += " (Tag {CASO})";
            _propertiesScrollPanel.Controls.Add(new Label { Text = lblCasoText, AutoSize = true, ForeColor = Color.DimGray });
            _txtPropTestCase = new RichTextBox { Width = ctrlWidth, Height = 40, Multiline = true, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 10) };
            _txtPropTestCase.TextChanged += (s, e) => {
                if (_isProgrammaticUpdate) return;
                _currentProjectData.TestCaseName = _txtPropTestCase.Text; _hasUnsavedChanges = true;
                _hostContext?.SetSessionMetadata("CurrentTestCaseName", _txtPropTestCase.Text); TriggerAutoSave();
            };
            var ctxTestCase = new ContextMenuStrip();
            ctxTestCase.Items.Add(LanguageManager.GetString("CtxPasteNoFormat"), null, (s, e) => { if (Clipboard.ContainsText()) _txtPropTestCase.SelectedText = Clipboard.GetText(TextDataFormat.Text); });
            _txtPropTestCase.ContextMenuStrip = ctxTestCase;
            _propertiesScrollPanel.Controls.Add(_txtPropTestCase);

            string lblQAText = LanguageManager.GetString("LblQA");
            if (!lblQAText.Contains("{QA}")) lblQAText += " (Tag {QA})";
            _propertiesScrollPanel.Controls.Add(new Label { Text = lblQAText, AutoSize = true, ForeColor = Color.DimGray });
            _txtPropQA = new TextBox { Width = ctrlWidth, Margin = new Padding(0, 0, 0, 2) };
            _txtPropQA.TextChanged += (s, e) => { if (_isProgrammaticUpdate) return; _currentProjectData.QAName = _txtPropQA.Text; _hasUnsavedChanges = true; TriggerAutoSave(); };
            _propertiesScrollPanel.Controls.Add(_txtPropQA);

            _chkPropDefaultQA = new CheckBox { Text = "Salvar como Padrão", AutoSize = true, ForeColor = Color.Gray, Font = new Font("Segoe UI", 8F), Margin = new Padding(0, 0, 0, 10), Checked = true };
            _chkPropDefaultQA.CheckedChanged += (s, e) => { SaveSettings(); };
            _propertiesScrollPanel.Controls.Add(_chkPropDefaultQA);

            string lblDateText = LanguageManager.GetString("LblDate");
            if (!lblDateText.Contains("{DATA}")) lblDateText += " (Tag {DATA})";
            _propertiesScrollPanel.Controls.Add(new Label { Text = lblDateText, AutoSize = true, ForeColor = Color.DimGray });
            _txtPropDate = new TextBox { Width = ctrlWidth, Margin = new Padding(0, 0, 0, 10) };
            _txtPropDate.TextChanged += (s, e) => { if (_isProgrammaticUpdate) return; _currentProjectData.TestDate = _txtPropDate.Text; _hasUnsavedChanges = true; TriggerAutoSave(); };
            _propertiesScrollPanel.Controls.Add(_txtPropDate);

            string lblObsText = LanguageManager.GetString("LblComments");
            if (!lblObsText.Contains("{OBS}")) lblObsText += " (Tag {OBS})";
            _propertiesScrollPanel.Controls.Add(new Label { Text = lblObsText, AutoSize = true, ForeColor = Color.DimGray });
            _txtPropComments = new RichTextBox { Width = ctrlWidth, Height = 55, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 0, 10) };
            _txtPropComments.TextChanged += (s, e) => { if (_isProgrammaticUpdate) return; _currentProjectData.Comments = _txtPropComments.Text; _hasUnsavedChanges = true; };
            _txtPropComments.Leave += (s, e) => { TriggerAutoSave(); };
            var ctxObs = new ContextMenuStrip();
            ctxObs.Items.Add(LanguageManager.GetString("CtxPasteFormat"), null, (s, e) => _txtPropComments.Paste());
            ctxObs.Items.Add(LanguageManager.GetString("CtxPasteNoFormat"), null, (s, e) => { if (Clipboard.ContainsText()) _txtPropComments.SelectedText = Clipboard.GetText(TextDataFormat.Text); });
            _txtPropComments.ContextMenuStrip = ctxObs;
            _propertiesScrollPanel.Controls.Add(_txtPropComments);

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblLayout"), AutoSize = true, ForeColor = Color.DimGray });
            _cmbPropLayout = new ComboBox { Width = ctrlWidth, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(0, 0, 0, 10), FlatStyle = FlatStyle.Flat };
            _cmbPropLayout.Items.AddRange(new object[] { "Padrão", "Mobile", "Compacto" });
            _propertiesScrollPanel.Controls.Add(_cmbPropLayout);

            _propertiesScrollPanel.Controls.Add(new Label { Text = LanguageManager.GetString("LblCols"), AutoSize = true, ForeColor = Color.DimGray });
            _numPropCols = new NumericUpDown { Width = ctrlWidth, Minimum = 1, Maximum = 3, Margin = new Padding(0, 0, 0, 10) };
            _propertiesScrollPanel.Controls.Add(_numPropCols);

            _cmbPropLayout.SelectedIndexChanged += (s, e) => {
                if (_isProgrammaticUpdate) return;
                _currentProjectData.ReportLayout = (LayoutMode)_cmbPropLayout.SelectedIndex;
                _numPropCols.Enabled = _cmbPropLayout.SelectedIndex == 1; // 1 é o Mobile
                _hasUnsavedChanges = true; TriggerAutoSave();
            };

            _numPropCols.ValueChanged += (s, e) => {
                if (_isProgrammaticUpdate) return;
                _currentProjectData.MobileColumns = (int)_numPropCols.Value;
                _hasUnsavedChanges = true; TriggerAutoSave();
            };

            _propertiesPanel.Controls.Add(_propertiesScrollPanel);
            _propertiesPanel.Controls.Add(pnlExport);

            _mainTable.Controls.Add(_propertiesPanel, 2, 0);
        }

        private void UpdatePropertiesPanelFromData()
        {
            if (_txtPropPrefix != null)
            {
                _isProgrammaticUpdate = true;

                _txtPropPrefix.Text = _currentProjectData.FilePrefix;
                _txtPropFileName.Text = _currentProjectData.FileName;
                _txtPropTestCase.Text = _currentProjectData.TestCaseName;
                _txtPropQA.Text = _currentProjectData.QAName;
                _txtPropDate.Text = string.IsNullOrEmpty(_currentProjectData.TestDate) ? DateTime.Now.ToString("dd/MM/yyyy") : _currentProjectData.TestDate;

                _cmbPropLayout.SelectedIndex = (int)_currentProjectData.ReportLayout;
                _numPropCols.Value = _currentProjectData.MobileColumns;
                _numPropCols.Enabled = _cmbPropLayout.SelectedIndex == 1;

                _txtPropComments.Text = _currentProjectData.Comments;

                _isProgrammaticUpdate = false;
            }
        }
    }
}