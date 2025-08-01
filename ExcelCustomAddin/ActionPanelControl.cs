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

            // Hiển thị file path tại toolStripFilePath
            UpdateFilePathDisplay();

            // Thiết lập ListView cho sheet - TẮT OwnerDraw để sử dụng background tự nhiên
            if (listofSheet != null)
            {
                listofSheet.OwnerDraw = false; // TẮT owner draw
                listofSheet.FullRowSelect = true;
                listofSheet.View = View.Details;
                listofSheet.HeaderStyle = ColumnHeaderStyle.None;
                listofSheet.HideSelection = false; // Đảm bảo selection hiển thị ngay cả khi ListView không có focus
                listofSheet.MultiSelect = false;   // Chỉ cho phép chọn một item

                // Tạo cột với chiều rộng ban đầu
                if (listofSheet.Columns.Count == 0)
                {
                    // Sử dụng chiều rộng mặc định nếu ListView chưa có kích thước
                    int initialWidth = listofSheet.Width > 0 ? listofSheet.Width - 4 : 400;
                    listofSheet.Columns.Add("Sheet", initialWidth);
                }

                // Đăng ký event để tự động điều chỉnh chiều rộng cột khi ListView thay đổi kích thước
                listofSheet.Resize += (sender, e) => UpdateColumnWidth();

                // Đăng ký event Load để cập nhật chiều rộng cột khi control được load
                this.Load += (sender, e) => UpdateColumnWidth();

                // Đăng ký event SizeChanged để đảm bảo cập nhật trong mọi trường hợp
                listofSheet.SizeChanged += (sender, e) => UpdateColumnWidth();

                // Đăng ký event để cập nhật context menu khi mở
                if (this.contextMenuStrip1 != null)
                {
                    this.contextMenuStrip1.Opening += ContextMenuStrip1_Opening;
                }
            }
        }

        /// <summary>
        /// Cập nhật chiều rộng cột để khớp với ListView
        /// </summary>
        private void UpdateColumnWidth()
        {
            if (listofSheet?.Columns != null && listofSheet.Columns.Count > 0 && listofSheet.Width > 0)
            {
                listofSheet.Columns[0].Width = listofSheet.Width - 4;
            }
        }

        /// <summary>
        /// Cập nhật hiển thị file path tại toolStripFilePath
        /// </summary>
        private void UpdateFilePathDisplay()
        {
            try
            {
                if (toolStripFilePath != null)
                {
                    // Lấy file path từ Excel Application
                    var app = Globals.ThisAddIn.Application;
                    if (app?.ActiveWorkbook != null)
                    {
                        string workbookPath = app.ActiveWorkbook.FullName;
                        if (!string.IsNullOrEmpty(workbookPath))
                        {
                            // Hiển thị đường dẫn đầy đủ
                            toolStripFilePath.Text = workbookPath;
                            toolStripFilePath.ToolTipText = workbookPath; // Tooltip để hiển thị full path
                        }
                        else
                        {
                            // Nếu workbook chưa được save
                            toolStripFilePath.Text = $"{app.ActiveWorkbook.Name} (Chưa lưu)";
                            toolStripFilePath.ToolTipText = "Workbook chưa được lưu";
                        }
                    }
                    else
                    {
                        toolStripFilePath.Text = "Không có file nào đang mở";
                        toolStripFilePath.ToolTipText = "";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in UpdateFilePathDisplay: {ex.Message}");
                if (toolStripFilePath != null)
                {
                    toolStripFilePath.Text = "Lỗi khi lấy đường dẫn file";
                    toolStripFilePath.ToolTipText = ex.Message;
                }
            }
        }

        /// <summary>
        /// Cập nhật context menu trước khi hiển thị
        /// </summary>
        private void ContextMenuStrip1_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                if (listofSheet?.SelectedItems != null && listofSheet.SelectedItems.Count > 0)
                {
                    var selectedItem = listofSheet.SelectedItems[0].Tag as ThisAddIn.SheetInfo;
                    if (selectedItem != null)
                    {
                        // Cập nhật text của menu item dựa trên trạng thái pin
                        if (this.btnPinSheet != null)
                        {
                            this.btnPinSheet.Text = selectedItem.IsPinned ? "Unpin Sheet" : "Pin Sheet";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ContextMenuStrip1_Opening: {ex.Message}");
                // Cancel context menu if error occurs
                e.Cancel = true;
            }
        }

        public event EventHandler FormatEvidenceEvent;
        public event EventHandler CreateEvidenceEvent;
        public event EventHandler FormatDocumentEvent;
        public event EventHandler ChangeSheetNameEvent;
        public event EventHandler InsertMultipleImagesEvent;

        public event EventHandler<PinSheetEventArgs> PinSheetEvent;

        /// <summary>
        /// Event args for pin sheet event
        /// </summary>
        public class PinSheetEventArgs : EventArgs
        {
            public string SheetName { get; set; }
            public bool IsPinned { get; set; }
        }

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

        /// <summary>
        /// btnChangeSheetName_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChangeSheetName_Click(object sender, EventArgs e)
        {
            if (this.ChangeSheetNameEvent != null)
                this.ChangeSheetNameEvent(this, e);
        }

        /// <summary>
        /// btnInsertPictures_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInsertPictures_Click(object sender, EventArgs e)
        {
            if (this.InsertMultipleImagesEvent != null)
                this.InsertMultipleImagesEvent(this, e);
        }

        /// <summary>
        /// btnPinSheet_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPinSheet_Click(object sender, EventArgs e)
        {
            try
            {
                if (listofSheet?.SelectedItems != null && listofSheet.SelectedItems.Count > 0)
                {
                    var selectedItem = listofSheet.SelectedItems[0].Tag as ThisAddIn.SheetInfo;
                    if (selectedItem != null && this.PinSheetEvent != null)
                    {
                        var args = new PinSheetEventArgs
                        {
                            SheetName = selectedItem.Name,
                            IsPinned = selectedItem.IsPinned
                        };
                        this.PinSheetEvent(this, args);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in btnPinSheet_Click: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"Có lỗi xảy ra khi thao tác với sheet: {ex.Message}", "Lỗi",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// BindSheetList
        /// </summary>
        /// <param name="sheets"></param>
        /// <param name="selectedSheetName"></param>
        public void BindSheetList(System.Collections.Generic.List<ThisAddIn.SheetInfo> sheets, string selectedSheetName = null)
        {
            try
            {
                if (listofSheet == null) return;

                listofSheet.Items.Clear();
                listofSheet.BeginUpdate();

                // Cập nhật chiều rộng cột trước khi bind data
                UpdateColumnWidth();

                // Cập nhật file path display
                UpdateFilePathDisplay();

                if (sheets != null)
                {
                    foreach (var sheet in sheets)
                    {
                        if (sheet == null) continue;

                        var lvi = new ListViewItem(sheet.Name ?? "Unknown");
                        lvi.Tag = sheet;

                        // Hiển thị trạng thái pin và màu sheet tab
                        string displayText = "";

                        // Thêm icon pin nếu sheet được pin
                        if (sheet.IsPinned)
                        {
                            displayText += "📌 "; // Pin icon
                        }

                        // Thêm màu sheet tab nếu có
                        if (sheet.HasTabColor)
                        {
                            displayText += "● "; // Bullet point để biểu thị có màu
                            lvi.ForeColor = sheet.TabColor; // Đặt màu text
                        }

                        displayText += sheet.Name;
                        lvi.Text = displayText;

                        listofSheet.Items.Add(lvi);

                        // Chọn item nếu tên sheet trùng với selectedSheetName
                        if (selectedSheetName != null && sheet.Name == selectedSheetName)
                        {
                            lvi.Selected = true;
                            lvi.Focused = true;
                            listofSheet.FocusedItem = lvi;
                        }
                    }
                }

                listofSheet.EndUpdate();

                // Đảm bảo item được chọn hiển thị trong viewport với kiểm tra an toàn
                try
                {
                    if (listofSheet.SelectedItems != null &&
                        listofSheet.SelectedItems.Count > 0 &&
                        listofSheet.SelectedItems[0].Index >= 0 &&
                        listofSheet.SelectedItems[0].Index < listofSheet.Items.Count)
                    {
                        listofSheet.EnsureVisible(listofSheet.SelectedItems[0].Index);
                    }
                }
                catch (Exception ex)
                {
                    // Log error but don't crash
                    System.Diagnostics.Debug.WriteLine($"Error in EnsureVisible: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in BindSheetList: {ex.Message}");
                // Ensure EndUpdate is called even if error occurs
                try
                {
                    listofSheet?.EndUpdate();
                }
                catch { }
            }
        }

        /// <summary>
        /// Cập nhật hiển thị file path từ bên ngoài
        /// </summary>
        public void RefreshFilePathDisplay()
        {
            UpdateFilePathDisplay();
        }
    }
}
