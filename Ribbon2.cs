using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;
using System.Net.Http;
using System.Net.Http.Json;
using Newtonsoft.Json;
using System.Threading.Tasks;



namespace OutlookAddIn1
{
    public partial class Ribbon2
    {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnCreateProfessional_Click(object sender, RibbonControlEventArgs e)
        {
           _ = GenerateAIMailAsync("formal");
        }

        private void btnCreateCasual_Click(object sender, RibbonControlEventArgs e)
        {
           _ = GenerateAIMailAsync("casual");
        }

        public static async Task<string> CallGeminiAsyncStatic(string prompt)
        {
            var apiKey = "AIzaSyD50oRMn2liwrw2DLm41uFBi5LILuUYCuE";
            var url = $"https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key={apiKey}";

            var requestBody = new
            {
                contents = new[]
                {
            new
            {
                parts = new[]
                {
                    new { text = prompt }
                }
            }
        }
            };

            var client = new HttpClient();

            int maxRetries = 3;
            int delayMs = 4000;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    var response = await client.PostAsJsonAsync(url, requestBody);

                    if (response.IsSuccessStatusCode)
                    {
                        var resultJson = await response.Content.ReadAsStringAsync();
                        dynamic result = JsonConvert.DeserializeObject(resultJson);
                        return result.candidates[0].content.parts[0].text;
                    }
                    else
                    {
                        var errorMsg = await response.Content.ReadAsStringAsync();

                        
                        if ((int)response.StatusCode == 503 && attempt < maxRetries)
                        {
                            await Task.Delay(delayMs);
                            continue;
                        }

                        MessageBox.Show($"Gemini API Error:\n{response.StatusCode}\n{errorMsg}", "Gemini Error");
                        return "Error: Unable to get response from GeminiAI.";
                    }
                }
                catch (Exception ex)
                {
                    if (attempt < maxRetries)
                    {
                        await Task.Delay(delayMs);
                        continue;
                    }

                    MessageBox.Show($"Gemini Exception:\n{ex.Message}", "Gemini Error");
                    return "Error";
                }
            }

            return "Error: Max retries exceeded.";
        }



        private async Task GenerateAIMailAsync(string tone)
        {
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector?.CurrentItem is Outlook.MailItem mailItem)

            {
                mailItem.Body = "";
                mailItem.Subject = "";
                mailItem.Save();


                using (var promptDialog = new PromptDialog("Create AI Mail", $"Enter what the email should be about (Tone: {tone})"))
                {
                    var result = promptDialog.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(promptDialog.UserInput))
                    {
                        string fullPrompt = $"Write an email in a {tone} tone. First, give a subject line in the format:\nSubject: <subject here>\nThen write the email body on a new line.\n\nTopic: {promptDialog.UserInput}";

                        string aiResponse = "";

                        using (var loading = new LoadingDialog("Generating AI-powered email..."))
                        {
                            loading.Show();
                            await Task.Delay(100);

                            aiResponse = await CallGeminiAsyncStatic(fullPrompt);

                            loading.Close();
                        }

                        if (aiResponse != "Error")
                        {
                      
                            string[] lines = aiResponse.Split(new[] { '\n' }, 2, StringSplitOptions.None);
                            string subjectLine = lines[0].Trim();
                            string bodyContent = lines.Length > 1 ? lines[1].Trim() : "";

                            if (subjectLine.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
                            {
                                mailItem.Subject = subjectLine.Substring(8).Trim();
                            }
                            else
                            {
                                mailItem.Subject = "[AI Generated Email]";
                            }

 
                            mailItem.Body = bodyContent;
                            mailItem.Save();

                            MessageBox.Show("AI-generated subject and body added.", "Success");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Prompt was empty or cancelled.");
                    }
                }
            }
            else
            {
                MessageBox.Show("No active mail window found.");
            }
        }

        private void btnSpellCheck_Click(object sender, RibbonControlEventArgs e)
        {
            _ = RunSpellCheckAsync();
        }

        private async Task RunSpellCheckAsync()
        {
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector?.CurrentItem is Outlook.MailItem mailItem)
            {
                string originalSubject = mailItem.Subject ?? "";
                string originalBody = mailItem.Body ?? "";

                if (string.IsNullOrWhiteSpace(originalSubject) && string.IsNullOrWhiteSpace(originalBody))
                {
                    MessageBox.Show("Both subject and body are empty.");
                    return;
                }

                string prompt = $"Check and correct any spelling mistakes in the following email.\n\nSubject:\n{originalSubject}\n\nBody:\n{originalBody}\n\nReturn only the corrected Subject and Body in the same format.";

                using (var loading = new LoadingDialog("Checking spelling..."))
                {
                    loading.Show();

                
                    await Task.Delay(100);

                    string correctedText = await CallGeminiAsyncStatic(prompt);

                    loading.Close();

                    if (correctedText != "Error")
                    {
                        string[] parts = correctedText.Split(new[] { "Body:" }, StringSplitOptions.None);
                        if (parts.Length == 2)
                        {
                            mailItem.Subject = parts[0].Replace("Subject:", "").Trim();
                            mailItem.Body = parts[1].Trim();
                            mailItem.Save();
                            MessageBox.Show("Spell-checked and corrected.", "Success");
                        }
                        else
                        {
                            MessageBox.Show("Unexpected AI response format. Couldn't update both fields.");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No active mail window found.");
            }
        }



        private async void btnlanguage_Click(object sender, RibbonControlEventArgs e)
        {
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector?.CurrentItem is Outlook.MailItem mailItem)
            {
                string subject = mailItem.Subject ?? "";
                string body = mailItem.Body ?? "";

                if (string.IsNullOrWhiteSpace(subject) && string.IsNullOrWhiteSpace(body))
                {
                    MessageBox.Show("Both subject and body are empty.");
                    return;
                }


                string emailBody = $"Subject:\n{subject}\n\nBody:\n{body}";
                using (var dialog = new TranslteDialog(emailBody))

                {
                    var result = dialog.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedLanguage))
                    {
                        string inputText = $"Subject:\n{subject}\n\nBody:\n{body}";
                        string prompt = $"Translate the following email into {dialog.SelectedLanguage}. Return in format:\nSubject: <translated subject>\nBody: <translated body>\n\nContent:\n{inputText}";

                        string aiResponse = "";

                       
                        using (var loading = new LoadingDialog("Translating email..."))
                        {
                            loading.Show();
                            Application.DoEvents();

                            aiResponse = await CallGeminiAsyncStatic(prompt);

                            loading.Close();
                        }

                        if (aiResponse != "Error" && !string.IsNullOrWhiteSpace(aiResponse))
                        {
                            
                            string[] parts = aiResponse.Split(new[] { "Body:" }, StringSplitOptions.None);
                            if (parts.Length == 2)
                            {
                                string translatedSubject = parts[0].Replace("Subject:", "").Trim();
                                string translatedBody = parts[1].Trim();

                                mailItem.Subject = translatedSubject;
                                mailItem.Body = translatedBody;
                            }
                            else
                            {
                              
                                mailItem.Body = aiResponse;
                            }

                            mailItem.Save();
                            MessageBox.Show($"Translated to {dialog.SelectedLanguage}.", "Success");
                        }
                        else
                        {
                            MessageBox.Show("Translation failed.");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No mail window found.");
            }
        }

        private void btnReplyAssist_Click(object sender, RibbonControlEventArgs e)
        {
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector?.CurrentItem is Outlook.MailItem mailItem)
            {
                using (var replyDialog = new ReplyDialog())
                {
                    var result = replyDialog.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(replyDialog.GeneratedReply))
                    {
                        string[] parts = replyDialog.GeneratedReply.Split(new[] { "Body:" }, StringSplitOptions.None);

                        if (parts.Length == 2)
                        {
                            string subjectPart = parts[0].Replace("Subject:", "").Trim();
                            string bodyPart = parts[1].Trim();

                            mailItem.Subject = subjectPart;
                            mailItem.Body = bodyPart;
                        }
                        else
                        {
                            mailItem.Body = replyDialog.GeneratedReply;
                        }

                        mailItem.Save();
                        MessageBox.Show("Reply email generated.", "Success");
                    }
                }
            }
            else
            {
                MessageBox.Show("No mail window open.");
            }
        }


    }
}
