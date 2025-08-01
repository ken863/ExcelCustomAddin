using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCustomAddin
{
    public partial class ActionPanelControl : UserControl
    {

        public ActionPanelControl()
        {
            InitializeComponent();

            // Thiết lập listbox để custom drawing
            if (listofSheet != null)
            {
                listofSheet.DrawMode = DrawMode.OwnerDrawFixed;
                listofSheet.DrawItem += ListofSheet_DrawItem;
                listofSheet.ItemHeight = 20; // Tăng chiều cao item để hiển thị màu tốt hơn
            }
        }

        /// <summary>
        /// Custom drawing cho listbox items với màu sheet
        /// </summary>
        private void ListofSheet_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            var listBox = sender as ListBox;
            var item = listBox.Items[e.Index];

            // Tùy chỉnh màu background thay vì sử dụng mặc định
            Color backgroundColor;
            Color textColor;

            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                // Màu khi item được chọn
                backgroundColor = Color.FromArgb(51, 153, 255); // Màu xanh dương nhẹ
                textColor = Color.White;
            }
            else if ((e.State & DrawItemState.HotLight) == DrawItemState.HotLight)
            {
                // Màu khi hover (nếu hỗ trợ)
                backgroundColor = Color.FromArgb(230, 240, 255); // Màu xanh rất nhẹ
                textColor = Color.Black;
            }
            else
            {
                // Màu bình thường - có thể tùy chỉnh theo sở thích
                backgroundColor = Color.FromArgb(248, 249, 250); // Màu xám rất nhẹ
                textColor = Color.Black;
            }

            // Vẽ background tùy chỉnh
            using (var backgroundBrush = new SolidBrush(backgroundColor))
            {
                e.Graphics.FillRectangle(backgroundBrush, e.Bounds);
            }

            if (item is ThisAddIn.SheetInfo sheetInfo)
            {
                // Vẽ màu tab của sheet (nếu có)
                if (sheetInfo.HasTabColor)
                {
                    using (var colorBrush = new SolidBrush(sheetInfo.TabColor))
                    {
                        var colorRect = new Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, 16, e.Bounds.Height - 4);
                        e.Graphics.FillRectangle(colorBrush, colorRect);
                        e.Graphics.DrawRectangle(Pens.Black, colorRect);
                    }
                }

                // Vẽ text tên sheet
                var textRect = new Rectangle(
                    e.Bounds.X + (sheetInfo.HasTabColor ? 22 : 4),
                    e.Bounds.Y,
                    e.Bounds.Width - (sheetInfo.HasTabColor ? 26 : 8),
                    e.Bounds.Height
                );

                using (var textBrush = new SolidBrush(textColor))
                {
                    var stringFormat = new StringFormat
                    {
                        Alignment = StringAlignment.Near,
                        LineAlignment = StringAlignment.Center
                    };

                    e.Graphics.DrawString(sheetInfo.Name, listBox.Font, textBrush, textRect, stringFormat);
                }
            }
            else
            {
                // Fallback cho các item không phải SheetInfo
                using (var textBrush = new SolidBrush(textColor))
                {
                    e.Graphics.DrawString(item.ToString(), listBox.Font, textBrush, e.Bounds);
                }
            }

            // Vẽ focus rectangle
            e.DrawFocusRectangle();
        }
        public event EventHandler FormatEvidenceEvent;
        public event EventHandler CreateEvidenceEvent;
        public event EventHandler FormatDocumentEvent;
        public event EventHandler ChangeSheetNameEvent;

        /// <summary>
        /// btnFormatEvidence_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFormatEvidence_Click(object sender, EventArgs e)
        {
            if (this.FormatEvidenceEvent != null)
                this.FormatEvidenceEvent(this, e);
        }

        /// <summary>
        /// btnCreateEvidence_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreateEvidence_Click(object sender, EventArgs e)
        {
            if (this.CreateEvidenceEvent != null)
                this.CreateEvidenceEvent(this, e);
        }

        /// <summary>
        /// btnFormatDocument_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFormatDocument_Click(object sender, EventArgs e)
        {
            if (this.FormatDocumentEvent != null)
                this.FormatDocumentEvent(this, e);
        }

        private void btnChangeSheetName_Click(object sender, EventArgs e)
        {
            if (this.ChangeSheetNameEvent != null)
                this.ChangeSheetNameEvent(this, e);
        }
    }
}
