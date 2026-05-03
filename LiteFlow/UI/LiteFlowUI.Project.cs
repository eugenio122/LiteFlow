using LiteFlow.Core;
using LiteFlow.Models;
using LiteFlow.Services;
using LiteTools.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LiteFlow.UI
{
    public partial class LiteFlowUI
    {
        private void TriggerAutoSave()
        {
            if (_isLoadingProject || !_isAutoSaveEnabled) return;
            _hasUnsavedChanges = true;
            if (!string.IsNullOrEmpty(_currentProjectPath))
            {
                _autoSaveTimer.Stop();
                _autoSaveTimer.Start();
            }
        }

        private void AutoSaveTimer_Tick(object? sender, EventArgs e)
        {
            _autoSaveTimer.Stop();
            if (_isSavingInBackground) { _autoSaveTimer.Start(); return; }
            SaveProjectInternalBackground(_currentProjectPath, true);
        }

        private void SaveProjectInternalBackground(string path, bool isAutoSave)
        {
            var projDataClone = new LiteFlowProjectData
            {
                TemplatePath = _currentProjectData.TemplatePath,
                FilePrefix = _currentProjectData.FilePrefix,
                FileName = _currentProjectData.FileName,
                TestCaseName = _currentProjectData.TestCaseName,
                QAName = _currentProjectData.QAName,
                TestDate = _currentProjectData.TestDate,
                Comments = _currentProjectData.Comments,
                ReportLayout = _currentProjectData.ReportLayout,
                MobileColumns = _currentProjectData.MobileColumns
            };

            var itemsToSave = new List<EvidenceData>();
            var pathsToRead = new List<string>();

            foreach (var item in GetItems())
            {
                itemsToSave.Add(new EvidenceData { StepId = item.StepId, Note = item.Note ?? "", TextBelowImage = item.TextBelowImage, IsEvidenceOnly = item.IsEvidenceOnly });
                pathsToRead.Add(item.DiskPath);
            }

            _isSavingInBackground = true;

            Task.Run(() => {
                string tempPath = path + ".tmp";
                try
                {
                    using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    using (var writer = new System.Text.Json.Utf8JsonWriter(fileStream))
                    {
                        writer.WriteStartObject();
                        writer.WriteString("TemplatePath", projDataClone.TemplatePath);
                        writer.WriteString("FilePrefix", projDataClone.FilePrefix);
                        writer.WriteString("FileName", projDataClone.FileName);
                        writer.WriteString("TestCaseName", projDataClone.TestCaseName);
                        writer.WriteString("QAName", projDataClone.QAName);
                        writer.WriteString("TestDate", projDataClone.TestDate);
                        writer.WriteString("Comments", projDataClone.Comments);
                        writer.WriteNumber("ReportLayout", (int)projDataClone.ReportLayout);
                        writer.WriteNumber("MobileColumns", projDataClone.MobileColumns);
                        writer.WritePropertyName("Steps");
                        writer.WriteStartArray();

                        for (int i = 0; i < itemsToSave.Count; i++)
                        {
                            writer.WriteStartObject();
                            writer.WriteString("StepId", itemsToSave[i].StepId);
                            string diskPath = pathsToRead[i];

                            if (!string.IsNullOrEmpty(diskPath) && File.Exists(diskPath))
                            {
                                // A OTIMIZAÇÃO CRÍTICA RECUPERADA: Zero alocação de Strings Gigantes na RAM
                                byte[] imageBytes = File.ReadAllBytes(diskPath);
                                writer.WriteBase64String("ImageDataBase64", imageBytes);
                                imageBytes = null!; // Liberta imediatamente para o GC
                            }
                            else
                            {
                                writer.WriteString("ImageDataBase64", "");
                            }

                            writer.WriteString("Note", itemsToSave[i].Note);
                            writer.WriteBoolean("TextBelowImage", itemsToSave[i].TextBelowImage);
                            writer.WriteBoolean("IsEvidenceOnly", itemsToSave[i].IsEvidenceOnly);
                            writer.WriteEndObject();
                        }
                        writer.WriteEndArray();
                        writer.WriteEndObject();
                    }

                    if (File.Exists(path)) File.Replace(tempPath, path, null, true);
                    else File.Move(tempPath, path);

                    this.BeginInvoke(new Action(() => {
                        _hasUnsavedChanges = false;
                        if (!isAutoSave) { UpdateAutoSaveUI(); UpdateProjectNameUI(); MessageBox.Show(LanguageManager.GetString("MsgProjectSaved"), LanguageManager.GetString("TitleLiteFlow"), MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    }));
                }
                catch (IOException) { if (!isAutoSave) this.BeginInvoke(new Action(() => { MessageBox.Show("O arquivo está temporariamente bloqueado, possivelmente devido à sincronização do OneDrive.\n\nAguarde o ícone de nuvem atualizar e tente salvar novamente.", "Arquivo Bloqueado", MessageBoxButtons.OK, MessageBoxIcon.Warning); })); }
                catch (Exception ex) { if (!isAutoSave) this.BeginInvoke(new Action(() => { MessageBox.Show($"Erro ao salvar: {ex.Message}", LanguageManager.GetString("TitleError"), MessageBoxButtons.OK, MessageBoxIcon.Error); })); }
                finally
                {
                    if (File.Exists(tempPath)) { try { File.Delete(tempPath); } catch { } }
                    _isSavingInBackground = false;
                    // Força uma varredura extrema de lixo após a escrita gigante
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
        }

        private void NewProject()
        {
            if (_hasUnsavedChanges && _historyRibbon.Controls.Count > 1)
            {
                var r = MessageBox.Show(LanguageManager.GetString("MsgSaveBeforeNew"), LanguageManager.GetString("TitleNew"), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (r == DialogResult.Cancel) return;
                if (r == DialogResult.Yes) SaveProjectCurrent();
            }

            if (_historyRibbon.Controls.Count > 1 && MessageBox.Show("Deseja realmente limpar a sessão e iniciar um novo teste?", "Aviso Crítico", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;

            _currentProjectData = new LiteFlowProjectData();
            _currentProjectData.ReportLayout = _defaultLayoutMode;
            _currentProjectData.MobileColumns = _defaultMobileColumns;
            _currentProjectData.QAName = _defaultQAName;
            _currentProjectData.FilePrefix = _defaultPrefix;
            _currentProjectData.TemplatePath = (!string.IsNullOrEmpty(_defaultTemplatePath) && File.Exists(_defaultTemplatePath) && MessageBox.Show(LanguageManager.GetString("MsgUseDefaultTemplate"), LanguageManager.GetString("TitleTemplate"), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) ? _defaultTemplatePath : "";

            ClearEvidenceHistory();
            _templateThumbnail.Invalidate();
            _currentProjectPath = "";
            _hasUnsavedChanges = false;
            _isAutoSaveEnabled = false;

            UpdatePropertiesPanelFromData();
            UpdateAutoSaveUI();
            UpdateProjectNameUI();
            FeedTestCaseToHostContext();

            _eventBus?.Publish(new SessionRestartedEvent());
        }

        private void RestartProject()
        {
            if (_historyRibbon.Controls.Count <= 1) return;

            var r = MessageBox.Show(
                "Deseja realmente apagar todas as capturas e recomeçar a execução?\nOs dados do painel lateral (Caso de Teste, Data, Executor) serão mantidos intactos.",
                "Recomeçar Projeto",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (r == DialogResult.Yes)
            {
                ClearEvidenceHistory();
                _hasUnsavedChanges = true;
                TriggerAutoSave();
                _eventBus?.Publish(new SessionRestartedEvent());
            }
        }

        private void SaveProjectCurrent()
        {
            if (string.IsNullOrEmpty(_currentProjectPath)) SaveProjectAs();
            else SaveProjectInternalBackground(_currentProjectPath, false);
        }

        private void SaveProjectAs()
        {
            string rawName = string.IsNullOrWhiteSpace(_currentProjectData.FileName) ? "Evidencias_LiteFlow" : $"{_currentProjectData.FilePrefix} {_currentProjectData.FileName}".Trim();
            string safeNameForWindows = SanitizeFileName(rawName);

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Projeto LiteFlow (*.lflow)|*.lflow", FileName = safeNameForWindows })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    _currentProjectPath = sfd.FileName;
                    string chosenName = Path.GetFileNameWithoutExtension(_currentProjectPath);
                    string prefix = _currentProjectData.FilePrefix ?? "";

                    if (!string.IsNullOrWhiteSpace(prefix) && chosenName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                        chosenName = chosenName.Substring(prefix.Length).Trim();

                    _currentProjectData.FileName = chosenName;

                    if (_txtPropFileName != null && !_txtPropFileName.IsDisposed)
                    {
                        _isProgrammaticUpdate = true;
                        _txtPropFileName.Text = _currentProjectData.FileName;
                        _isProgrammaticUpdate = false;
                    }

                    SaveProjectInternalBackground(_currentProjectPath, false);
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

                            if (string.IsNullOrWhiteSpace(_currentProjectData.FileName)) _currentProjectData.FileName = Path.GetFileNameWithoutExtension(_currentProjectPath);

                            _templateThumbnail.Invalidate();
                            _isLoadingProject = true;
                            foreach (var step in data.Steps)
                            {
                                string path = Path.Combine(_sessionTempDir, $"img_{Guid.NewGuid():N}.png");
                                File.WriteAllBytes(path, Convert.FromBase64String(step.ImageDataBase64));
                                step.ImageDataBase64 = null!;
                                AddToHistoryFromDisk(path, null, step.Note, step.TextBelowImage, step.IsEvidenceOnly, step.StepId);
                            }

                            data.Steps.Clear();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();

                            _isLoadingProject = false;
                            _hasUnsavedChanges = false;
                            _isAutoSaveEnabled = true;

                            UpdatePropertiesPanelFromData();
                            UpdateAutoSaveUI();
                            UpdateProjectNameUI();
                            FeedTestCaseToHostContext();
                        }
                    }
                    catch (IOException ex) { _isLoadingProject = false; MessageBox.Show($"Não foi possível carregar o projeto. O arquivo pode estar bloqueado pelo OneDrive ou outro processo de sincronização.\n\nAguarde o ícone de nuvem atualizar e tente novamente.\n\nDetalhe: {ex.Message}", "Arquivo Bloqueado", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    catch (Exception ex) { _isLoadingProject = false; MessageBox.Show(string.Format(LanguageManager.GetString("MsgError"), ex.Message), LanguageManager.GetString("TitleError")); }
                }
            }
        }

        private string SanitizeFileName(string name)
        {
            string invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            foreach (char c in invalidChars) name = name.Replace(c.ToString(), "");
            return name;
        }

        private Dictionary<string, string> BuildExportTags()
        {
            string dataExport = _currentProjectData.TestDate;

            if (string.IsNullOrWhiteSpace(dataExport))
            {
                dataExport = (_txtPropDate != null && !string.IsNullOrWhiteSpace(_txtPropDate.Text))
                             ? _txtPropDate.Text
                             : DateTime.Now.ToString("dd/MM/yyyy");

                _currentProjectData.TestDate = dataExport;
            }

            return new Dictionary<string, string>
            {
                { "{CASO}", _currentProjectData.TestCaseName ?? "" },
                { "{QA}", _currentProjectData.QAName ?? "" },
                { "{DATA}", dataExport },
                { "{OBS}", _currentProjectData.Comments ?? "" }
            };
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
                        _currentProjectData.TemplatePath = destPath;

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

            string templateToUse = !string.IsNullOrEmpty(_currentProjectData.TemplatePath) && File.Exists(_currentProjectData.TemplatePath)
                                    ? _currentProjectData.TemplatePath
                                    : _defaultTemplatePath;

            if (string.IsNullOrEmpty(templateToUse) || !File.Exists(templateToUse))
            {
                MessageBox.Show(LanguageManager.GetString("MsgExportNoTemplate"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            _currentProjectData.TemplatePath = templateToUse;

            SaveSettings();
            TriggerAutoSave();

            string rawName = string.IsNullOrWhiteSpace(_currentProjectData.FileName) ? "Evidencias_LiteFlow" : $"{_currentProjectData.FilePrefix} {_currentProjectData.FileName}".Trim();
            string safeFileName = SanitizeFileName(rawName);

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Word (*.docx)|*.docx", FileName = safeFileName + ".docx" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        ExportService.ExportToWord(_currentProjectData, itemsToExport, sfd.FileName, BuildExportTags());
                        this.Cursor = Cursors.Default;
                        MessageBox.Show(string.Format(LanguageManager.GetString("MsgExportSuccessWord"), sfd.FileName), LanguageManager.GetString("TitleSuccess"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) { this.Cursor = Cursors.Default; MessageBox.Show(string.Format(LanguageManager.GetString("MsgError"), ex.Message), LanguageManager.GetString("TitleError"), MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }

        private void BtnApplyToPdf_Click(object? sender, EventArgs e)
        {
            var itemsToExport = GetItems();
            if (itemsToExport.Count == 0) return;

            string templateToUse = !string.IsNullOrEmpty(_currentProjectData.TemplatePath) && File.Exists(_currentProjectData.TemplatePath)
                                    ? _currentProjectData.TemplatePath
                                    : _defaultTemplatePath;

            if (string.IsNullOrEmpty(templateToUse) || !File.Exists(templateToUse))
            {
                MessageBox.Show(LanguageManager.GetString("MsgExportNoTemplate"), LanguageManager.GetString("Warning"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            _currentProjectData.TemplatePath = templateToUse;

            SaveSettings();
            TriggerAutoSave();

            string rawName = string.IsNullOrWhiteSpace(_currentProjectData.FileName) ? "Evidencias_LiteFlow" : $"{_currentProjectData.FilePrefix} {_currentProjectData.FileName}".Trim();
            string safeFileName = SanitizeFileName(rawName);

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", FileName = safeFileName + ".pdf" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        ExportService.ExportToPdf(_currentProjectData, itemsToExport, sfd.FileName, BuildExportTags());
                        this.Cursor = Cursors.Default;
                        MessageBox.Show(LanguageManager.GetString("MsgExportSuccessPdf"), LanguageManager.GetString("TitleLiteFlowPdf"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) { this.Cursor = Cursors.Default; MessageBox.Show(string.Format(LanguageManager.GetString("MsgErrorPdf"), ex.Message), LanguageManager.GetString("TitleError"), MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
            }
        }
    }
}