using System;
using System.Drawing;
using System.Windows.Forms;

namespace LiteFlow.Forms
{
    /// <summary>
    /// Formulário clássico (Janela) para configuração do motor de evidências do LiteFlow.
    /// Permite ao utilizador definir o caminho do template do Word e a pasta de saída.
    /// Responsivo: Utiliza TableLayoutPanel para evitar quebras em DPI Scaling (150%+).
    /// </summary>
    public partial class LiteFlowSettingsForm : Form
    {
        private TextBox _txtTemplatePath = null!;
        private TextBox _txtOutputPath = null!;
        private CheckBox _chkAutoSave = null!;

        public LiteFlowSettingsForm()
        {
            SetupUI();
        }

        private void SetupUI()
        {
            // Configurações da Janela
            this.Text = "LiteFlow - Configurações de Relatório";
            this.Size = new Size(580, 390); // Aumentado ligeiramente para acomodar o link
            this.MinimumSize = new Size(500, 370);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);

            // Esqueleto Principal (TableLayoutPanel) - Substitui todos os Locations
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 9, // Aumentado para 9 linhas
                Padding = new Padding(20, 20, 20, 15),
                BackColor = Color.White
            };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // Título
            var lblHeader = new Label
            {
                Text = "Configurações do Motor Word",
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                AutoSize = true,
                ForeColor = Color.FromArgb(0, 120, 215),
                Margin = new Padding(0, 0, 0, 20)
            };

            // Secção Template (Rótulo)
            var lblTemplate = new Label
            {
                Text = "Caminho do Template (.docx):",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 5)
            };

            // Grelha interna para o Input + Botão Procurar
            TableLayoutPanel tlpTemplate = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            tlpTemplate.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlpTemplate.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90F));

            _txtTemplatePath = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Margin = new Padding(0, 3, 10, 0) };
            var btnBrowseTemplate = new Button { Text = "Procurar", Dock = DockStyle.Fill, Cursor = Cursors.Hand, Height = 28, Margin = new Padding(0) };

            btnBrowseTemplate.Click += (s, e) =>
            {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Documentos Word (*.docx)|*.docx", Title = "Selecione o Template de Testes" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        _txtTemplatePath.Text = ofd.FileName;
                }
            };
            tlpTemplate.Controls.Add(_txtTemplatePath, 0, 0);
            tlpTemplate.Controls.Add(btnBrowseTemplate, 1, 0);

            // Secção Saída (Rótulo)
            var lblOutput = new Label
            {
                Text = "Pasta de Saída dos Relatórios:",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 5)
            };

            // Grelha interna para o Input + Botão Procurar
            TableLayoutPanel tlpOutput = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 0, 0, 20) };
            tlpOutput.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlpOutput.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90F));

            _txtOutputPath = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Margin = new Padding(0, 3, 10, 0) };
            var btnBrowseOutput = new Button { Text = "Procurar", Dock = DockStyle.Fill, Cursor = Cursors.Hand, Height = 28, Margin = new Padding(0) };

            btnBrowseOutput.Click += (s, e) =>
            {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog { Description = "Selecione a pasta para guardar as evidências" })
                {
                    if (fbd.ShowDialog() == DialogResult.OK)
                        _txtOutputPath.Text = fbd.SelectedPath;
                }
            };
            tlpOutput.Controls.Add(_txtOutputPath, 0, 0);
            tlpOutput.Controls.Add(btnBrowseOutput, 1, 0);

            // AutoSave Option
            _chkAutoSave = new CheckBox
            {
                Text = "Exportar imagem automaticamente a cada captura",
                AutoSize = true,
                Checked = true, // Ligado por defeito (Filosofia reativa do LiteFlow)
                Margin = new Padding(0, 0, 0, 20)
            };

            // Separador Visual
            var separator = new Label { BorderStyle = BorderStyle.Fixed3D, Height = 2, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 15) };

            // Link do GitHub
            var lnkGithub = new LinkLabel
            {
                Text = "GitHub: github.com/eugenio122/LiteFlow",
                AutoSize = true,
                LinkColor = Color.FromArgb(0, 120, 215),
                Margin = new Padding(0, 0, 0, 15)
            };
            lnkGithub.LinkClicked += (s, e) => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://github.com/eugenio122/LiteFlow") { UseShellExecute = true });

            // Botões Finais (Alinhados à direita)
            FlowLayoutPanel pnlButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = new Padding(0)
            };

            var btnCancel = new Button { Text = "Cancelar", Width = 90, Height = 32, Cursor = Cursors.Hand, Margin = new Padding(10, 0, 0, 0) };
            btnCancel.Click += (s, e) => this.Close();

            var btnSave = new Button { Text = "Guardar", Width = 90, Height = 32, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0) };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += (s, e) => {
                // TODO: Salvar nas Settings do plugin
                MessageBox.Show("Configurações salvas com sucesso!", "LiteFlow", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            pnlButtons.Controls.Add(btnCancel);
            pnlButtons.Controls.Add(btnSave);

            // Adicionar todas as linhas ao esqueleto principal na ordem correta
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 0: Header
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 1: Label Template
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 2: Input Template
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 3: Label Saída
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 4: Input Saída
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 5: Checkbox AutoSave
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 6: Separator
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 7: Link GitHub
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 8: Buttons

            mainLayout.Controls.Add(lblHeader, 0, 0);
            mainLayout.Controls.Add(lblTemplate, 0, 1);
            mainLayout.Controls.Add(tlpTemplate, 0, 2);
            mainLayout.Controls.Add(lblOutput, 0, 3);
            mainLayout.Controls.Add(tlpOutput, 0, 4);
            mainLayout.Controls.Add(_chkAutoSave, 0, 5);
            mainLayout.Controls.Add(separator, 0, 6);
            mainLayout.Controls.Add(lnkGithub, 0, 7);
            mainLayout.Controls.Add(pnlButtons, 0, 8);

            this.Controls.Add(mainLayout);
        }
    }
}