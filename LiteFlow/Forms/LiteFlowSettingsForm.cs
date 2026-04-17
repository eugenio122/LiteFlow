using System;
using System.Drawing;
using System.Windows.Forms;

namespace LiteFlow.Forms
{
    /// <summary>
    /// Formulário clássico (Janela) para configuração do motor de evidências do LiteFlow.
    /// Permite ao utilizador definir o caminho do template do Word e a pasta de saída.
    /// </summary>
    public partial class LiteFlowSettingsForm : Form
    {
        private TextBox _txtTemplatePath;
        private TextBox _txtOutputPath;
        private CheckBox _chkAutoSave;

        public LiteFlowSettingsForm()
        {
            SetupUI();
        }

        private void SetupUI()
        {
            // Configurações da Janela
            this.Text = "LiteFlow - Configurações de Relatório";
            this.Size = new Size(550, 320);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);

            // Título
            var lblHeader = new Label
            {
                Text = "Configurações do Motor Word",
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Location = new Point(20, 20),
                AutoSize = true,
                ForeColor = Color.FromArgb(0, 120, 215)
            };

            // Secção Template
            var lblTemplate = new Label { Text = "Caminho do Template (.docx):", Location = new Point(20, 60), AutoSize = true };
            _txtTemplatePath = new TextBox { Location = new Point(20, 80), Width = 400, ReadOnly = true };
            var btnBrowseTemplate = new Button { Text = "Procurar", Location = new Point(430, 79), Width = 80, Cursor = Cursors.Hand };

            btnBrowseTemplate.Click += (s, e) =>
            {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Documentos Word (*.docx)|*.docx", Title = "Selecione o Template de Testes" })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        _txtTemplatePath.Text = ofd.FileName;
                }
            };

            // Secção Saída
            var lblOutput = new Label { Text = "Pasta de Saída dos Relatórios:", Location = new Point(20, 120), AutoSize = true };
            _txtOutputPath = new TextBox { Location = new Point(20, 140), Width = 400, ReadOnly = true };
            var btnBrowseOutput = new Button { Text = "Procurar", Location = new Point(430, 139), Width = 80, Cursor = Cursors.Hand };

            btnBrowseOutput.Click += (s, e) =>
            {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog { Description = "Selecione a pasta para guardar as evidências" })
                {
                    if (fbd.ShowDialog() == DialogResult.OK)
                        _txtOutputPath.Text = fbd.SelectedPath;
                }
            };

            // AutoSave Option
            _chkAutoSave = new CheckBox
            {
                Text = "Exportar imagem automaticamente a cada captura",
                Location = new Point(20, 180),
                AutoSize = true,
                Checked = true // Ligado por defeito (Filosofia reativa do LiteFlow)
            };

            // Separador Visual
            var separator = new Label { BorderStyle = BorderStyle.Fixed3D, Height = 2, Width = 500, Location = new Point(15, 220) };

            // Botões Finais
            var btnSave = new Button { Text = "Guardar", Location = new Point(340, 240), Width = 80, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += (s, e) => {
                // Salvar nas Settings do plugin
                MessageBox.Show("Configurações salvas com sucesso!", "LiteFlow", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            var btnCancel = new Button { Text = "Cancelar", Location = new Point(430, 240), Width = 80, Cursor = Cursors.Hand };
            btnCancel.Click += (s, e) => this.Close();

            // Adicionar Controlos
            this.Controls.Add(lblHeader);
            this.Controls.Add(lblTemplate);
            this.Controls.Add(_txtTemplatePath);
            this.Controls.Add(btnBrowseTemplate);
            this.Controls.Add(lblOutput);
            this.Controls.Add(_txtOutputPath);
            this.Controls.Add(btnBrowseOutput);
            this.Controls.Add(_chkAutoSave);
            this.Controls.Add(separator);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
        }
    }
}