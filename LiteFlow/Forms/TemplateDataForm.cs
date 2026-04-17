using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using LiteFlow.Models;

namespace LiteFlow.Forms
{
    public partial class TemplateDataForm : Form
    {
        private TextBox _txtPrefix = null!;
        private CheckBox _chkDefaultPrefix = null!;

        private TextBox _txtFileName = null!;
        private RichTextBox _txtCaso = null!;
        private RichTextBox _txtObs = null!;

        private TextBox _txtQA = null!;
        private CheckBox _chkDefaultQA = null!;
        private TextBox _txtData = null!;

        private ComboBox _cmbLayoutMode = null!;
        private NumericUpDown _numColumns = null!;

        public bool SaveQADefault => _chkDefaultQA.Checked;
        public bool SavePrefixDefault => _chkDefaultPrefix.Checked;

        public string SuggestedFileName => $"{_txtPrefix.Text} {_txtFileName.Text}".Trim();

        public TemplateDataForm(string defaultQA, string defaultPrefix, LiteFlowProjectData projectData)
        {
            SetupUI(defaultQA, defaultPrefix, projectData);
        }

        private void SetupUI(string defaultQA, string defaultPrefix, LiteFlowProjectData projectData)
        {
            this.Text = "Informações do Relatório";
            this.Size = new Size(440, 540);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9F);

            var lblTitle = new Label { Text = "Preencha os dados do Teste", Font = new Font("Segoe UI", 12F, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };

            var lblPrefix = new Label { Text = "Prefixo:", Location = new Point(20, 60), AutoSize = true };
            _txtPrefix = new TextBox { Location = new Point(20, 80), Width = 120, Text = !string.IsNullOrWhiteSpace(projectData.FilePrefix) ? projectData.FilePrefix : defaultPrefix };

            var lblFileName = new Label { Text = "Nome do Arquivo:", Location = new Point(150, 60), AutoSize = true };
            _txtFileName = new TextBox { Location = new Point(150, 80), Width = 250, Text = projectData.FileName };
            _txtFileName.KeyDown += LiteFlowUI.ForcePlainTextPaste;

            _chkDefaultPrefix = new CheckBox { Text = "Salvar Prefixo como Padrão", Location = new Point(20, 110), AutoSize = true, ForeColor = Color.DimGray, Checked = true };

            // AGORA COM MULTILINE PARA ACEITAR LINKS DO JIRA SEM CORTAR
            var lblCaso = new Label { Text = "Cenário / Caso de Teste: (Tag {CASO})", Location = new Point(20, 140), AutoSize = true };
            _txtCaso = new RichTextBox { Location = new Point(20, 160), Width = 380, Height = 45, Multiline = true, Text = projectData.TestCaseName, BorderStyle = BorderStyle.FixedSingle };
            var ctxCaso = new ContextMenuStrip();
            ctxCaso.Items.Add("Colar sem formatação", null, (s, e) => { if (Clipboard.ContainsText()) _txtCaso.SelectedText = Clipboard.GetText(TextDataFormat.Text); });
            _txtCaso.ContextMenuStrip = ctxCaso;

            var lblQA = new Label { Text = "Executor (QA): (Tag {QA})", Location = new Point(20, 215), AutoSize = true };
            _txtQA = new TextBox { Location = new Point(20, 235), Width = 160, Text = !string.IsNullOrWhiteSpace(projectData.QAName) ? projectData.QAName : defaultQA };

            var lblData = new Label { Text = "Data: (Tag {DATA})", Location = new Point(240, 215), AutoSize = true };
            _txtData = new TextBox { Location = new Point(240, 235), Width = 160, Text = !string.IsNullOrWhiteSpace(projectData.TestDate) ? projectData.TestDate : DateTime.Now.ToString("dd/MM/yyyy") };

            _chkDefaultQA = new CheckBox { Text = "Salvar QA como Padrão", Location = new Point(20, 265), AutoSize = true, ForeColor = Color.DimGray, Checked = true };

            var pnlLayout = new GroupBox { Text = "Configuração Visual", Location = new Point(20, 295), Size = new Size(380, 80), ForeColor = Color.FromArgb(0, 120, 215) };

            var lblMode = new Label { Text = "Modo:", Location = new Point(15, 25), AutoSize = true, ForeColor = Color.Black };
            _cmbLayoutMode = new ComboBox { Location = new Point(15, 45), Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbLayoutMode.Items.AddRange(new object[] { "Padrão", "Mobile", "Compacto" });
            _cmbLayoutMode.SelectedIndex = (int)projectData.ReportLayout;

            var lblCols = new Label { Text = "Colunas:", Location = new Point(125, 25), AutoSize = true, ForeColor = Color.Black };
            _numColumns = new NumericUpDown { Location = new Point(125, 45), Width = 50, Minimum = 1, Maximum = 3, Value = projectData.MobileColumns };
            _numColumns.Enabled = _cmbLayoutMode.SelectedIndex == 1;

            _cmbLayoutMode.SelectedIndexChanged += (s, e) => { _numColumns.Enabled = _cmbLayoutMode.SelectedIndex == 1; };

            pnlLayout.Controls.AddRange(new Control[] { lblMode, _cmbLayoutMode, lblCols, _numColumns });

            var lblObs = new Label { Text = "Comentários: (Tag {OBS})", Location = new Point(20, 385), AutoSize = true };
            _txtObs = new RichTextBox { Location = new Point(20, 405), Width = 380, Multiline = true, Height = 50, Text = projectData.Comments, BorderStyle = BorderStyle.FixedSingle };
            var ctxObs = new ContextMenuStrip();
            ctxObs.Items.Add("Colar (Manter Formatação) (Ctrl+V)", null, (s, e) => _txtObs.Paste());
            ctxObs.Items.Add("Colar sem formatação", null, (s, e) => { if (Clipboard.ContainsText()) _txtObs.SelectedText = Clipboard.GetText(TextDataFormat.Text); });
            _txtObs.ContextMenuStrip = ctxObs;

            var btnExport = new Button { Text = "Confirmar Exportação", Location = new Point(240, 465), Width = 160, Height = 35, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnExport.FlatAppearance.BorderSize = 0;
            btnExport.Click += (s, e) => { this.DialogResult = DialogResult.OK; this.Close(); };

            this.Controls.AddRange(new Control[] {
                lblTitle, lblPrefix, _txtPrefix, lblFileName, _txtFileName, _chkDefaultPrefix,
                lblCaso, _txtCaso, lblQA, _txtQA, lblData, _txtData, _chkDefaultQA,
                pnlLayout, lblObs, _txtObs, btnExport
            });
        }

        public void UpdateProjectData(LiteFlowProjectData data)
        {
            data.FilePrefix = _txtPrefix.Text;
            data.FileName = _txtFileName.Text;
            data.TestCaseName = _txtCaso.Text;
            data.QAName = _txtQA.Text;
            data.TestDate = _txtData.Text;
            data.Comments = _txtObs.Text;

            data.ReportLayout = (LayoutMode)_cmbLayoutMode.SelectedIndex;
            data.MobileColumns = (int)_numColumns.Value;
        }

        public Dictionary<string, string> GetTags()
        {
            return new Dictionary<string, string>
            {
                { "{CASO}", _txtCaso.Text },
                { "{QA}", _txtQA.Text },
                { "{DATA}", _txtData.Text },
                { "{OBS}", _txtObs.Text }
            };
        }
    }
}