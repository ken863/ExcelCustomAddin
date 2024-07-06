using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace ExcelCustomAddin
{
    public partial class ActionPanelControl : UserControl
    {

        public ActionPanelControl()
        {
            InitializeComponent();
        }

        public event EventHandler TranslateClickEvent;
        public event EventHandler TranslateDoEvent;
        public event EventHandler TranslateCompletedEvent;

        private void ButtonTranslate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSourceText.Text.Trim()))
            {
                return;
            }

            this.UpdateProgressBar(true);
            this.UpdateButton(false);
            bgwTranslate.RunWorkerAsync(txtSourceText.Text.Trim());
        }

        /// <summary>
        /// bgwTranslate_DoWork
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void bgwTranslate_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                var apiKey = "sk-proj-KHBw6jj2cKclN3xmD5olT3BlbkFJekvhNIP9ykw0F1xIScCD";
                var chatGPTClient = new ChatGPTClient(apiKey);

                var text = e.Argument.ToString();
                if (!string.IsNullOrEmpty(text))
                {
                    var response = await chatGPTClient.CallChatGPTAsync(text);
                    this.UpdateText(response);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.UpdateProgressBar(false);
                this.UpdateButton(true);
            }
        }

        /// <summary>
        /// UpdateText
        /// </summary>
        /// <param name="text"></param>
        private void UpdateText(string text)
        {
            if (txtDesText.InvokeRequired)
            {
                // Nếu cần phải gọi Invoke, sử dụng phương thức này để gọi hàm từ thread khác
                txtDesText.Invoke(new Action<string>(UpdateText), text);
            }
            else
            {
                // Nếu không cần phải gọi Invoke, cập nhật trực tiếp
                txtDesText.Text = text;
            }
        }

        /// <summary>
        /// UpdateProgressBar
        /// </summary>
        /// <param name="isVisible"></param>
        private void UpdateProgressBar(bool isVisible)
        {
            if (progressBar.InvokeRequired)
            {
                // Nếu cần phải gọi Invoke, sử dụng phương thức này để gọi hàm từ thread khác
                progressBar.Invoke(new Action<bool>(UpdateProgressBar), isVisible);
            }
            else
            {
                // Nếu không cần phải gọi Invoke, cập nhật trực tiếp
                progressBar.Visible = isVisible;
            }
        }

        /// <summary>
        /// UpdateButton
        /// </summary>
        /// <param name="isVisible"></param>
        private void UpdateButton(bool isEnable)
        {
            if (progressBar.InvokeRequired)
            {
                // Nếu cần phải gọi Invoke, sử dụng phương thức này để gọi hàm từ thread khác
                buttonTranslate.Invoke(new Action<bool>(UpdateButton), isEnable);
            }
            else
            {
                // Nếu không cần phải gọi Invoke, cập nhật trực tiếp
                buttonTranslate.Enabled = isEnable;
            }
        }

        private void listofSheet_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
