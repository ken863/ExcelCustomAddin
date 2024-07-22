namespace ExcelCustomAddin
{
    using Newtonsoft.Json.Linq;
    using System;
    using System.Net;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;

    public class ChatGPTClient
    {
        public ChatGPTClient()
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }

        public async Task<string> CallChatGPTAsync(string prompt)
        {
            if (string.IsNullOrEmpty(Properties.Settings.Default.API_KEY) || string.IsNullOrEmpty(Properties.Settings.Default.MODEL))
            {
                return "Vui lòng setting API Key!!!";
            }

            try
            {
                string _apiKey = Properties.Settings.Default.API_KEY;
                HttpClient client = new HttpClient();
                var url = "https://api.openai.com/v1/chat/completions";
                var payload = new
                {
                    model = Properties.Settings.Default.MODEL,
                    messages = new[] {
                        new {
                            role = "system", content = $"Dịch sang tiếng việt: {prompt}"
                        }
                    }
                };

                var content = new StringContent(JObject.FromObject(payload).ToString(), Encoding.UTF8, "application/json");
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}");
                var response = await client.PostAsync(url, content);

                if (response.IsSuccessStatusCode)
                {
                    //It would be better to make sure this request actually made it through
                    JObject responseObject = JObject.Parse(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                    string result = responseObject["choices"]?[0]?["message"]?["content"]?.ToString();
                    return result;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return string.Empty;
        }
    }
}
