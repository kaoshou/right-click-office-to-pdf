using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace OfficePdf
{
    internal static class Drawing
    {
        public static int Pixels(Graphics graphics, int logical)
        {
            return (int)Math.Round(logical * graphics.DpiX / 96f);
        }

        public static void FillPill(Graphics graphics, Rectangle bounds, Color color)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0) return;
            int d = Math.Min(bounds.Height, bounds.Width);
            using (GraphicsPath path = new GraphicsPath())
            using (SolidBrush brush = new SolidBrush(color))
            {
                path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
                path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
                path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
                path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                graphics.FillPath(brush, path);
            }
        }
    }

    public class FileList : ListView
    {
        private const TextFormatFlags TextFlags = TextFormatFlags.SingleLine |
            TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix |
            TextFormatFlags.NoPadding | TextFormatFlags.PreserveGraphicsClipping;
        private ImageList rowHeightImages;

        public FileList()
        {
            View = View.Details;
            OwnerDraw = true;
            DoubleBuffered = true;
            BorderStyle = BorderStyle.None;
            HeaderStyle = ColumnHeaderStyle.Nonclickable;
            FullRowSelect = true;
            HideSelection = false;
            MultiSelect = false;
            ShowItemToolTips = true;
            BackColor = Color.White;
            ForeColor = Color.FromArgb(32, 46, 62);
        }

        private void UpdateRowHeight()
        {
            if (!IsHandleCreated || Disposing || IsDisposed) return;
            using (Graphics g = CreateGraphics())
            {
                int height = Math.Min(256, Math.Max(Drawing.Pixels(g, 38),
                    Font.Height + Drawing.Pixels(g, 12)));
                if (rowHeightImages != null && rowHeightImages.ImageSize.Height == height) return;
                // Native report-view row height follows the small image list, including
                // when there are no images. Keep native item, keyboard and UIA behavior.
                ImageList previous = rowHeightImages;
                rowHeightImages = new ImageList();
                rowHeightImages.ImageSize = new Size(1, height);
                SmallImageList = rowHeightImages;
                if (previous != null) previous.Dispose();
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            UpdateRowHeight();
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            UpdateRowHeight();
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == 0x02E3) UpdateRowHeight(); // WM_DPICHANGED_AFTERPARENT
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && rowHeightImages != null)
            {
                SmallImageList = null;
                rowHeightImages.Dispose();
                rowHeightImages = null;
            }
            base.Dispose(disposing);
        }

        protected override void OnDrawColumnHeader(DrawListViewColumnHeaderEventArgs e)
        {
            bool contrast = SystemInformation.HighContrast;
            Color background = contrast ? SystemColors.Control : Color.FromArgb(243, 246, 250);
            Color foreground = contrast ? SystemColors.ControlText : Color.FromArgb(73, 88, 106);
            using (SolidBrush brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, e.Bounds);
            DrawSeparator(e.Graphics, e.Bounds, contrast);
            DrawText(e.Graphics, e.Header.Text, e.Bounds, foreground);
        }

        protected override void OnDrawItem(DrawListViewItemEventArgs e)
        {
            // Each subitem paints its own background. Native hover events sometimes
            // request the item alone, so painting over it here would erase its text.
            if (View != View.Details) e.DrawDefault = true;
        }

        private static void DrawSeparator(Graphics g, Rectangle bounds, bool contrast)
        {
            using (Pen pen = new Pen(contrast ? SystemColors.ControlDark : Color.FromArgb(229, 234, 240)))
                g.DrawLine(pen, bounds.Left, bounds.Bottom - 1, bounds.Right, bounds.Bottom - 1);
        }

        private void DrawText(Graphics g, string text, Rectangle cell, Color color)
        {
            int padding = Drawing.Pixels(g, 12);
            cell.Inflate(-padding, 0);
            if (cell.Width > 0) TextRenderer.DrawText(g, text, Font, cell, color, TextFlags);
        }

        protected override void OnDrawSubItem(DrawListViewSubItemEventArgs e)
        {
            Rectangle cell = e.Bounds;
            if (cell.Width <= 0 || cell.Height <= 0) return;
            bool contrast = SystemInformation.HighContrast;
            bool selected = e.Item.Selected && (Focused || !HideSelection);
            Color background = contrast ? SystemColors.Window : BackColor;
            Color text = contrast ? SystemColors.WindowText : ForeColor;
            if (selected)
            {
                background = contrast ? SystemColors.Highlight : (Focused ?
                    Color.FromArgb(232, 242, 255) : Color.FromArgb(240, 244, 249));
                if (contrast) text = SystemColors.HighlightText;
            }
            using (SolidBrush brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, cell);
            DrawSeparator(e.Graphics, cell, contrast);
            if (e.ColumnIndex == 1 && !contrast)
                DrawStatus(e.Graphics, cell, e.SubItem.Text);
            else
                DrawText(e.Graphics, e.SubItem.Text, cell, text);

            if (e.ColumnIndex == 0 && e.Item.Focused && Focused && ShowFocusCues)
            {
                cell.Inflate(-1, -2);
                if (cell.Width > 0 && cell.Height > 0)
                    ControlPaint.DrawFocusRectangle(e.Graphics, cell, text, background);
            }
        }

        private void DrawStatus(Graphics g, Rectangle cell, string state)
        {
            Color ink, tint;
            switch (state)
            {
                case "Converting":
                case "轉換中": ink = Color.FromArgb(21, 95, 197); tint = Color.FromArgb(225, 238, 255); break;
                case "Completed":
                case "已完成": ink = Color.FromArgb(8, 119, 94); tint = Color.FromArgb(224, 244, 237); break;
                case "Failed":
                case "未完成": ink = Color.FromArgb(174, 52, 52); tint = Color.FromArgb(255, 233, 232); break;
                default: ink = Color.FromArgb(82, 97, 114); tint = Color.FromArgb(234, 239, 244); break;
            }
            int padding = Drawing.Pixels(g, 12);
            int inset = Drawing.Pixels(g, 8);
            TextFormatFlags measure = TextFormatFlags.SingleLine | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix;
            int width = TextRenderer.MeasureText(g, state, Font, Size.Empty, measure).Width + 2 * inset;
            int height = Math.Min(cell.Height - 2, Font.Height + Drawing.Pixels(g, 6));
            Rectangle pill = new Rectangle(cell.Left + padding, cell.Top + (cell.Height - height) / 2,
                Math.Min(cell.Width - padding - 2, width), height);
            if (pill.Width <= 0 || pill.Height <= 0) return;
            SmoothingMode original = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Drawing.FillPill(g, pill, tint);
            g.SmoothingMode = original;
            pill.Inflate(-Math.Min(inset, pill.Width / 4), 0);
            TextRenderer.DrawText(g, state, Font, pill, ink, TextFlags | TextFormatFlags.HorizontalCenter);
        }
    }

    public class BatchProgress : ProgressBar
    {
        public BatchProgress()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            // Standard ProgressBar properties and accessibility still use native
            // progress messages. Repaint our track after that native state changes.
            if (m.Msg == 0x401 || m.Msg == 0x402 || m.Msg == 0x403 ||
                m.Msg == 0x405 || m.Msg == 0x406) Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Color background = Parent == null ? SystemColors.Control : Parent.BackColor;
            e.Graphics.Clear(background);
            if (ClientSize.Width <= 0 || ClientSize.Height <= 0) return;
            int height = Math.Min(ClientSize.Height, Math.Max(1, Drawing.Pixels(e.Graphics, 7)));
            Rectangle track = new Rectangle(0, (ClientSize.Height - height) / 2, ClientSize.Width, height);
            bool contrast = SystemInformation.HighContrast;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Drawing.FillPill(e.Graphics, track, contrast ? SystemColors.ControlDark : Color.FromArgb(223, 230, 239));
            if (Maximum <= Minimum || Value <= Minimum) return;
            double fraction = ((double)Value - Minimum) / ((double)Maximum - Minimum);
            track.Width = Math.Max(1, Math.Min(track.Width, (int)Math.Round(track.Width * fraction)));
            Drawing.FillPill(e.Graphics, track, contrast ? SystemColors.Highlight : Color.FromArgb(21, 95, 197));
        }
    }
}
