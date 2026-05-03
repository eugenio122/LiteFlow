using LiteFlow.Core;
using LiteFlow.Models;
using LiteTools.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LiteFlow.UI
{
    public partial class LiteFlowUI
    {
        private void SetupHistoryRibbon()
        {
            _historyWrapper = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };

            _historyHeaderPanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, Padding = new Padding(5, 5, 5, 10), WrapContents = false };
            _lblHistoryTitle = new Label { Text = "⏱️ Histórico", AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Segoe UI", 8.5F, FontStyle.Bold), Margin = new Padding(0, 0, 0, 10) };

            FlowLayoutPanel pnlButtons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Margin = new Padding(0) };

            _btnPaste = new Button { Text = "📋 Colar Imagem", Width = 130, Height = 28, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8F), Cursor = Cursors.Hand, BackColor = Color.WhiteSmoke, Margin = new Padding(5, 0, 0, 5), TextAlign = ContentAlignment.MiddleLeft };
            _btnPaste.FlatAppearance.BorderSize = 0;
            _btnPaste.Click += (s, e) => {
                if (!PasteImageFromClipboard()) MessageBox.Show(LanguageManager.GetString("MsgNoImageClipboard"), LanguageManager.GetString("TitlePaste"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            _btnAddBlank = new Button { Text = "➕ Tela em Branco", Width = 130, Height = 28, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8F), Cursor = Cursors.Hand, BackColor = Color.WhiteSmoke, Margin = new Padding(5, 0, 0, 0), TextAlign = ContentAlignment.MiddleLeft };
            _btnAddBlank.FlatAppearance.BorderSize = 0;
            _btnAddBlank.Click += (s, e) => {
                Bitmap blank = new Bitmap(1024, 768);
                using (Graphics g = Graphics.FromImage(blank)) { g.Clear(Color.White); }
                Task.Run(() => {
                    string path = Path.Combine(_sessionTempDir, $"img_{Guid.NewGuid():N}.png");
                    blank.Save(path, ImageFormat.Png);
                    this.BeginInvoke(new Action(() => { AddToHistoryFromDisk(path, blank); }));
                });
            };

            pnlButtons.Controls.Add(_btnPaste);
            pnlButtons.Controls.Add(_btnAddBlank);

            _historyHeaderPanel.Controls.Add(_lblHistoryTitle);
            _historyHeaderPanel.Controls.Add(pnlButtons);

            _historyRibbon = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, WrapContents = false, FlowDirection = FlowDirection.TopDown, BackColor = Color.FromArgb(245, 245, 245), AllowDrop = true, Padding = new Padding(5) };
            _historyRibbon.DragEnter += (s, e) => { if (e.Data.GetDataPresent(typeof(PictureBox))) e.Effect = DragDropEffects.Move; };
            _historyRibbon.DragDrop += HistoryRibbon_DragDrop;

            _templateThumbnail = new PictureBox { Width = 110, Height = 70, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.WhiteSmoke, Padding = new Padding(2), Margin = new Padding(5), Cursor = Cursors.Hand };
            _templateThumbnail.Paint += TemplateThumbnail_Paint;
            _templateThumbnail.Click += (s, e) => {
                string tPath = !string.IsNullOrEmpty(_currentProjectData.TemplatePath) ? _currentProjectData.TemplatePath : _defaultTemplatePath;
                if (File.Exists(tPath)) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(tPath) { UseShellExecute = true });
                else MessageBox.Show(LanguageManager.GetString("MsgNoTemplate"), LanguageManager.GetString("TitleTemplate"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            _historyRibbon.Controls.Add(_templateThumbnail);

            _historyWrapper.Controls.Add(_historyRibbon);
            _historyWrapper.Controls.Add(_historyHeaderPanel);

            _mainTable.Controls.Add(_historyWrapper, 0, 0);
        }

        private void TemplateThumbnail_Paint(object? sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            Color borderColor = _isDarkMode ? Color.FromArgb(80, 80, 80) : Color.LightGray;
            e.Graphics.DrawRectangle(new Pen(borderColor), 0, 0, _templateThumbnail.Width - 1, _templateThumbnail.Height - 1);

            string tPath = !string.IsNullOrEmpty(_currentProjectData.TemplatePath) ? _currentProjectData.TemplatePath : _defaultTemplatePath;
            bool hasTemplate = File.Exists(tPath);

            Color textColor = _isDarkMode ? Color.WhiteSmoke : Color.Black;

            using (StringFormat topFormat = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Near })
            using (Font f2 = new Font("Segoe UI", 7, FontStyle.Regular))
            using (Brush textBrush = new SolidBrush(textColor))
            {
                string name = hasTemplate ? Path.GetFileName(tPath) : LanguageManager.GetString("BlankDoc");
                if (name.Length > 18) name = name.Substring(0, 15) + "...";

                e.Graphics.DrawString(name, f2, textBrush, new Rectangle(0, 4, _templateThumbnail.Width, 18), topFormat);
            }

            int docWidth = 34;
            int docHeight = 40;
            int docX = (_templateThumbnail.Width - docWidth) / 2;
            int docY = 24;

            e.Graphics.FillRectangle(Brushes.White, docX, docY, docWidth, docHeight);
            e.Graphics.DrawRectangle(Pens.SteelBlue, docX, docY, docWidth, docHeight);
            e.Graphics.FillRectangle(Brushes.SteelBlue, docX, docY, docWidth, 12);

            using (StringFormat centerFormat = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            using (Font f = new Font("Segoe UI", 6, FontStyle.Bold))
            {
                e.Graphics.DrawString("WORD", f, Brushes.White, new Rectangle(docX, docY, docWidth, 12), centerFormat);
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
            _evidenceDiskPaths.Clear(); // <-- LIMPO DE FORMA SEGURA. Sem a variável de Lista antiga a causar embaraços.

            if (_editorCore != null) { _editorCore.LoadImage(new Bitmap(10, 10)); }

            while (_pendingImages.Count > 0) { _pendingImages.Dequeue()?.Dispose(); }

            SelectEvidence(null);

            try
            {
                if (Directory.Exists(_sessionTempDir)) Directory.Delete(_sessionTempDir, true);
                Directory.CreateDirectory(_sessionTempDir);
            }
            catch { }

            GC.Collect();
            GC.WaitForPendingFinalizers();
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
            else if (isUndoRedo)
            {
                _eventBus?.Publish(new StepRestoredEvent(item.StepId));
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

            _eventBus?.Publish(new StepDeletedEvent(item.StepId));

            ReindexHistory(); TriggerAutoSave();
        }

        private void PublishReorderedSteps()
        {
            var ids = GetItems().Select(i => i.StepId).ToList();
            _eventBus?.Publish(new StepsReorderedEvent(ids));
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
            ReindexHistory();
            TriggerAutoSave();
            PublishReorderedSteps();
        }

        private void Thumbnail_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is PictureBox pb && pb.Tag is EvidenceItem item)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                string indexText = "-";

                if (!item.IsEvidenceOnly)
                {
                    int logicalIndex = 1;
                    foreach (Control c in _historyRibbon.Controls)
                    {
                        if (c == pb) break;
                        if (c is PictureBox otherPb && otherPb.Tag is EvidenceItem otherItem && !otherItem.IsEvidenceOnly) logicalIndex++;
                    }
                    indexText = logicalIndex.ToString();
                }

                e.Graphics.FillEllipse(new SolidBrush(Color.FromArgb(200, 0, 120, 215)), 4, 4, 18, 18);
                e.Graphics.DrawString(indexText, new Font("Segoe UI", 8, FontStyle.Bold), Brushes.White, new Point(7, 5));
                e.Graphics.FillEllipse(new SolidBrush(Color.FromArgb(220, 220, 53, 69)), new Rectangle(pb.Width - 24, 4, 18, 18));
                e.Graphics.DrawString("X", new Font("Segoe UI", 7, FontStyle.Bold), Brushes.White, new Point(pb.Width - 19, 6));

                if (item.IsEvidenceOnly)
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(220, 255, 140, 0)), new Rectangle(2, pb.Height - 16, 85, 14));
                    e.Graphics.DrawString($"👁️ {LanguageManager.GetString("EvidenceOnly")}", new Font("Segoe UI", 6, FontStyle.Bold), Brushes.White, new Point(3, pb.Height - 15));
                }
            }
        }

        private void Thumbnail_MouseDown(object? sender, MouseEventArgs e)
        {
            if (sender is PictureBox pb && pb.Tag is EvidenceItem item)
            {
                if (e.X >= pb.Width - 25 && e.Y <= 25)
                {
                    RemoveEvidenceFromUI(item, false);
                    return;
                }

                if (_stepNoteTextBox != null && _stepNoteTextBox.Focused) TriggerAutoSave();

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

        private List<EvidenceItem> GetItems()
        {
            var list = new List<EvidenceItem>();
            foreach (PictureBox pb in _historyRibbon.Controls)
            {
                if (pb.Tag is EvidenceItem item) list.Add(item);
            }
            return list;
        }
    }
}