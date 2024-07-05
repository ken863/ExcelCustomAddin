using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            bgwTranslate.RunWorkerAsync(txtSourceText.Text.Trim());
            this.UpdateProgressBar(true);
        }

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
        }

        private void bgwTranslate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.UpdateProgressBar(false);
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
    }
}
