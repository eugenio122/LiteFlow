using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using LiteFlow.Models;

namespace LiteFlow.Controller
{
    public class ImageEditorCore
    {
        private PictureBox _canvas;
        private TextBox _floatingTextBox;

        public EditorTool CurrentTool { get; set; } = EditorTool.Arrow;
        public Color CurrentColor { get; set; } = Color.Red;
        public int CurrentThickness { get; set; } = 4;
        public string CurrentFontFamily { get; set; } = "Segoe UI";
        public int CurrentFontSize { get; set; } = 14;

        public Bitmap? WorkingImage { get; private set; }
        private Stack<Bitmap> _undoStack = new Stack<Bitmap>();
        private Stack<Bitmap> _redoStack = new Stack<Bitmap>();

        public bool CanUndo => _undoStack.Count > 0;
        public bool CanRedo => _redoStack.Count > 0;

        public Action? OnImageEdited;

        private bool _isDrawing = false;
        private Point _imageStart;
        private Point _imageCurrent;
        private Point _screenStart;
        private Point _screenCurrent;

        private Rectangle _cropRect;
        private int _activeResizeHandle = -1;
        private const int HANDLE_SIZE = 8;

        public ImageEditorCore(PictureBox canvas, TextBox floatingTextBox)
        {
            _canvas = canvas;
            _floatingTextBox = floatingTextBox;

            _canvas.MouseDown += Canvas_MouseDown;
            _canvas.MouseMove += Canvas_MouseMove;
            _canvas.MouseUp += Canvas_MouseUp;
            _canvas.Paint += Canvas_Paint;

            _floatingTextBox.KeyDown += FloatingTextBox_KeyDown;
            _floatingTextBox.Leave += FloatingTextBox_Leave;
        }

        // ======================================================================
        // Limpeza Destrutiva das Pilhas (Mata o vazamento de 1GB RAM)
        // ======================================================================
        private void ClearStacks()
        {
            while (_undoStack.Count > 0) _undoStack.Pop()?.Dispose();
            while (_redoStack.Count > 0) _redoStack.Pop()?.Dispose();
        }

        private void ClearRedoStack()
        {
            while (_redoStack.Count > 0) _redoStack.Pop()?.Dispose();
        }
        // ======================================================================

        public void LoadImage(Bitmap img)
        {
            ClearStacks(); // Limpa a memória lixo do GDI+ antes de carregar nova imagem
            WorkingImage?.Dispose();
            WorkingImage = new Bitmap(img);
            _canvas.Image = WorkingImage;
            CancelCurrentAction();
        }

        public void Undo()
        {
            if (_undoStack.Count > 0)
            {
                _redoStack.Push(new Bitmap(WorkingImage!));
                WorkingImage?.Dispose();
                WorkingImage = _undoStack.Pop();
                _canvas.Image = WorkingImage;
                OnImageEdited?.Invoke();
                _canvas.Invalidate();
            }
        }

        public void Redo()
        {
            if (_redoStack.Count > 0)
            {
                _undoStack.Push(new Bitmap(WorkingImage!));
                WorkingImage?.Dispose();
                WorkingImage = _redoStack.Pop();
                _canvas.Image = WorkingImage;
                OnImageEdited?.Invoke();
                _canvas.Invalidate();
            }
        }

        public void CancelCurrentAction()
        {
            _floatingTextBox.Visible = false;
            _cropRect = Rectangle.Empty;
            _canvas.Invalidate();
        }

        private Point GetImageCoords(Point mousePoint, bool clamp = true)
        {
            if (_canvas.Image == null || WorkingImage == null) return mousePoint;
            float ratio = Math.Min((float)_canvas.Width / WorkingImage.Width, (float)_canvas.Height / WorkingImage.Height);
            int dw = (int)(WorkingImage.Width * ratio); int dh = (int)(WorkingImage.Height * ratio);
            int ox = (_canvas.Width - dw) / 2; int oy = (_canvas.Height - dh) / 2;
            int x = (int)((mousePoint.X - ox) / ratio);
            int y = (int)((mousePoint.Y - oy) / ratio);

            if (clamp)
            {
                x = Math.Max(0, Math.Min(x, WorkingImage.Width - 1));
                y = Math.Max(0, Math.Min(y, WorkingImage.Height - 1));
            }
            return new Point(x, y);
        }

        private Rectangle GetImageScreenBounds()
        {
            if (WorkingImage == null) return Rectangle.Empty;
            float ratio = Math.Min((float)_canvas.Width / WorkingImage.Width, (float)_canvas.Height / WorkingImage.Height);
            int dw = (int)(WorkingImage.Width * ratio); int dh = (int)(WorkingImage.Height * ratio);
            int ox = (_canvas.Width - dw) / 2; int oy = (_canvas.Height - dh) / 2;
            return new Rectangle(ox, oy, dw, dh);
        }

        private Rectangle[] GetSelectionHandles(Rectangle bounds)
        {
            int hs = HANDLE_SIZE;
            return new Rectangle[] {
                new Rectangle(bounds.Left - hs/2, bounds.Top - hs/2, hs, hs),
                new Rectangle(bounds.Left + bounds.Width/2 - hs/2, bounds.Top - hs/2, hs, hs),
                new Rectangle(bounds.Right - hs/2, bounds.Top - hs/2, hs, hs),
                new Rectangle(bounds.Right - hs/2, bounds.Top + bounds.Height/2 - hs/2, hs, hs),
                new Rectangle(bounds.Right - hs/2, bounds.Bottom - hs/2, hs, hs),
                new Rectangle(bounds.Left + bounds.Width/2 - hs/2, bounds.Bottom - hs/2, hs, hs),
                new Rectangle(bounds.Left - hs/2, bounds.Bottom - hs/2, hs, hs),
                new Rectangle(bounds.Left - hs/2, bounds.Top + bounds.Height/2 - hs/2, hs, hs)
            };
        }

        private void Canvas_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || WorkingImage == null) return;

            if (CurrentTool == EditorTool.Select)
            {
                var handles = GetSelectionHandles(GetImageScreenBounds());
                for (int i = 0; i < handles.Length; i++)
                {
                    if (handles[i].Contains(e.Location) && (i == 3 || i == 4 || i == 5))
                    {
                        _activeResizeHandle = i; _isDrawing = true;
                        _undoStack.Push(new Bitmap(WorkingImage));
                        ClearRedoStack(); // GC Fix
                        return;
                    }
                }
            }

            if (CurrentTool == EditorTool.Text)
            {
                if (_floatingTextBox.Visible) { CommitText(); return; }
                _undoStack.Push(new Bitmap(WorkingImage));
                ClearRedoStack(); // GC Fix
                _imageStart = GetImageCoords(e.Location);
                _floatingTextBox.Font = new Font(CurrentFontFamily, CurrentFontSize, FontStyle.Bold);
                _floatingTextBox.ForeColor = CurrentColor;
                _floatingTextBox.Text = "";
                _floatingTextBox.Location = e.Location;
                _floatingTextBox.Size = new Size(200, 30);
                _floatingTextBox.Visible = true;
                _floatingTextBox.Focus();
                return;
            }

            _isDrawing = true;
            _screenStart = e.Location;
            _imageStart = GetImageCoords(e.Location);
            _undoStack.Push(new Bitmap(WorkingImage));
            ClearRedoStack(); // GC Fix
        }

        private void Canvas_MouseMove(object? sender, MouseEventArgs e)
        {
            if (CurrentTool == EditorTool.Select && !_isDrawing)
            {
                var handles = GetSelectionHandles(GetImageScreenBounds());
                bool onHandle = false;
                for (int i = 0; i < handles.Length; i++)
                {
                    if (handles[i].Contains(e.Location) && (i == 3 || i == 4 || i == 5)) { onHandle = true; break; }
                }
                _canvas.Cursor = onHandle ? Cursors.SizeAll : Cursors.Default;
            }

            if (!_isDrawing || WorkingImage == null) return;

            _screenCurrent = e.Location;
            _imageCurrent = GetImageCoords(e.Location, _activeResizeHandle == -1);

            if (_activeResizeHandle != -1) { _canvas.Invalidate(); return; }

            if (CurrentTool == EditorTool.Pen || CurrentTool == EditorTool.Highlight)
            {
                using (Graphics g = Graphics.FromImage(WorkingImage))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    Color c = CurrentTool == EditorTool.Highlight ? Color.FromArgb(100, CurrentColor) : CurrentColor;
                    using (Pen p = new Pen(c, CurrentThickness) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                    {
                        g.DrawLine(p, _imageStart, _imageCurrent);
                    }
                }
                _imageStart = _imageCurrent;
                _canvas.Invalidate();
            }
            else { _canvas.Invalidate(); }
        }

        private void Canvas_MouseUp(object? sender, MouseEventArgs e)
        {
            if (!_isDrawing || WorkingImage == null) return;
            _isDrawing = false;

            if (_activeResizeHandle != -1)
            {
                int newW = WorkingImage.Width;
                int newH = WorkingImage.Height;
                if (_activeResizeHandle == 3 || _activeResizeHandle == 4) newW = Math.Max(10, _imageCurrent.X);
                if (_activeResizeHandle == 5 || _activeResizeHandle == 4) newH = Math.Max(10, _imageCurrent.Y);

                Bitmap resized = new Bitmap(newW, newH);
                using (Graphics g = Graphics.FromImage(resized))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.Clear(Color.White);
                    g.DrawImage(WorkingImage, 0, 0, newW, newH);
                }
                WorkingImage.Dispose(); WorkingImage = resized; _canvas.Image = WorkingImage;
                OnImageEdited?.Invoke();

                _activeResizeHandle = -1; _canvas.Invalidate(); return;
            }

            if (CurrentTool == EditorTool.Crop) { _canvas.Invalidate(); return; }

            if (CurrentTool != EditorTool.Pen && CurrentTool != EditorTool.Highlight && CurrentTool != EditorTool.Text)
            {
                using (Graphics g = Graphics.FromImage(WorkingImage))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    using (Pen p = new Pen(CurrentColor, CurrentThickness))
                    {
                        int x = Math.Min(_imageStart.X, _imageCurrent.X); int y = Math.Min(_imageStart.Y, _imageCurrent.Y);
                        int w = Math.Abs(_imageStart.X - _imageCurrent.X); int h = Math.Abs(_imageStart.Y - _imageCurrent.Y);

                        if (CurrentTool == EditorTool.Line) g.DrawLine(p, _imageStart, _imageCurrent);
                        else if (CurrentTool == EditorTool.Arrow) { p.CustomEndCap = new AdjustableArrowCap(5, 5); g.DrawLine(p, _imageStart, _imageCurrent); }
                        else if (CurrentTool == EditorTool.Shape) g.DrawRectangle(p, x, y, w, h);
                    }
                }
            }
            OnImageEdited?.Invoke(); _canvas.Invalidate();
        }

        private void Canvas_Paint(object? sender, PaintEventArgs e)
        {
            if (WorkingImage == null) return;

            if (CurrentTool == EditorTool.Select && _activeResizeHandle == -1)
            {
                var bounds = GetImageScreenBounds();
                using (Pen dashPen = new Pen(Color.Black, 1) { DashStyle = DashStyle.Dash })
                {
                    e.Graphics.DrawRectangle(dashPen, bounds);
                    var handles = GetSelectionHandles(bounds);
                    foreach (var i in new[] { 3, 4, 5 })
                    {
                        e.Graphics.FillRectangle(Brushes.White, handles[i]);
                        e.Graphics.DrawRectangle(Pens.Black, handles[i]);
                    }
                }
            }

            if (_activeResizeHandle != -1)
            {
                int newW = WorkingImage.Width;
                int newH = WorkingImage.Height;
                if (_activeResizeHandle == 3 || _activeResizeHandle == 4) newW = Math.Max(10, _imageCurrent.X);
                if (_activeResizeHandle == 5 || _activeResizeHandle == 4) newH = Math.Max(10, _imageCurrent.Y);

                float ratio = Math.Min((float)_canvas.Width / WorkingImage.Width, (float)_canvas.Height / WorkingImage.Height);
                int ox = (_canvas.Width - (int)(WorkingImage.Width * ratio)) / 2;
                int oy = (_canvas.Height - (int)(WorkingImage.Height * ratio)) / 2;

                using (Pen p = new Pen(Color.Black, 1) { DashStyle = DashStyle.Dash })
                    e.Graphics.DrawRectangle(p, ox, oy, (int)(newW * ratio), (int)(newH * ratio));
            }

            if (_isDrawing && CurrentTool == EditorTool.Crop)
            {
                int x = Math.Min(_screenStart.X, _screenCurrent.X); int y = Math.Min(_screenStart.Y, _screenCurrent.Y);
                int w = Math.Abs(_screenStart.X - _screenCurrent.X); int h = Math.Abs(_screenStart.Y - _screenCurrent.Y);
                _cropRect = new Rectangle(x, y, w, h);
                using (Pen p = new Pen(Color.Red, 2) { DashStyle = DashStyle.Dash }) e.Graphics.DrawRectangle(p, _cropRect);
            }
            else if (_isDrawing && CurrentTool != EditorTool.Pen && CurrentTool != EditorTool.Highlight && CurrentTool != EditorTool.Select && CurrentTool != EditorTool.Text)
            {
                using (Pen p = new Pen(CurrentColor, CurrentThickness))
                {
                    int x = Math.Min(_screenStart.X, _screenCurrent.X); int y = Math.Min(_screenStart.Y, _screenCurrent.Y);
                    int w = Math.Abs(_screenStart.X - _screenCurrent.X); int h = Math.Abs(_screenStart.Y - _screenCurrent.Y);
                    if (CurrentTool == EditorTool.Line) e.Graphics.DrawLine(p, _screenStart, _screenCurrent);
                    else if (CurrentTool == EditorTool.Arrow) { p.CustomEndCap = new AdjustableArrowCap(5, 5); e.Graphics.DrawLine(p, _screenStart, _screenCurrent); }
                    else if (CurrentTool == EditorTool.Shape) e.Graphics.DrawRectangle(p, x, y, w, h);
                }
            }
        }

        public void ConfirmCrop()
        {
            if (_cropRect == Rectangle.Empty || WorkingImage == null) return;
            Point p1 = GetImageCoords(new Point(_cropRect.Left, _cropRect.Top));
            Point p2 = GetImageCoords(new Point(_cropRect.Right, _cropRect.Bottom));
            int x = Math.Max(0, p1.X); int y = Math.Max(0, p1.Y);
            int w = Math.Min(WorkingImage.Width - x, p2.X - p1.X); int h = Math.Min(WorkingImage.Height - y, p2.Y - p1.Y);
            if (w <= 0 || h <= 0) { CancelCurrentAction(); return; }

            _undoStack.Push(new Bitmap(WorkingImage));
            ClearRedoStack(); // GC Fix

            Bitmap cropped = new Bitmap(w, h);
            using (Graphics g = Graphics.FromImage(cropped)) { g.DrawImage(WorkingImage, new Rectangle(0, 0, w, h), new Rectangle(x, y, w, h), GraphicsUnit.Pixel); }
            WorkingImage.Dispose(); WorkingImage = cropped; _canvas.Image = WorkingImage;
            CancelCurrentAction(); OnImageEdited?.Invoke();
        }

        private void FloatingTextBox_KeyDown(object? sender, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) CancelCurrentAction(); }

        private void FloatingTextBox_Leave(object? sender, EventArgs e) { CommitText(); }

        private void CommitText()
        {
            if (!_floatingTextBox.Visible || WorkingImage == null) return;
            if (!string.IsNullOrWhiteSpace(_floatingTextBox.Text))
            {
                using (Graphics g = Graphics.FromImage(WorkingImage))
                {
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    using (Brush b = new SolidBrush(CurrentColor))
                    using (Font f = new Font(CurrentFontFamily, CurrentFontSize, FontStyle.Bold))
                        g.DrawString(_floatingTextBox.Text, f, b, _imageStart);
                }
                OnImageEdited?.Invoke();
            }
            else
            {
                var imgToDispose = _undoStack.Pop();
                imgToDispose?.Dispose(); // GC Fix: Descarta a cópia se não escreveu nada
            }
            CancelCurrentAction();
        }
    }
}