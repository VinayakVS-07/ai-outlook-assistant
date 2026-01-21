using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class TranslteDialog : Form
    {
        public string SelectedLanguage { get; private set; }
        public string TranslatedText { get; private set; }

        private ComboBox comboLanguages;
        private Button btnTranslate;
        private Button btnApply;
        private Button btnCancel;
        private TextBox txtPreview;

        private string originalText;

        public TranslteDialog(string emailBody)
        {
            InitializeComponent();
            this.originalText = emailBody;

            this.Text = "Translate Email";
            this.Size = new Size(700, 500);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            Label lblLang = new Label() { Text = "Select Language:", Top = 20, Left = 20, Width = 120 };
            comboLanguages = new ComboBox() { Left = 150, Top = 18, Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            comboLanguages.Items.AddRange(GetLanguageList().ToArray());

            btnTranslate = new Button() { Text = "Preview Translation", Left = 370, Top = 16, Width = 150 };
            btnTranslate.Click += BtnTranslate_Click;

            txtPreview = new TextBox()
            {
                Left = 20,
                Top = 60,
                Width = 640,
                Height = 300,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Segoe UI", 10)
            };

            btnApply = new Button() { Text = "Apply", Left = 480, Top = 380, Width = 80, DialogResult = DialogResult.OK };
            btnApply.Click += (s, e) => { this.TranslatedText = txtPreview.Text; this.SelectedLanguage = comboLanguages.Text; this.Close(); };

            btnCancel = new Button() { Text = "Cancel", Left = 580, Top = 380, Width = 80, DialogResult = DialogResult.Cancel };

            this.Controls.Add(lblLang);
            this.Controls.Add(comboLanguages);
            this.Controls.Add(btnTranslate);
            this.Controls.Add(txtPreview);
            this.Controls.Add(btnApply);
            this.Controls.Add(btnCancel);
        }
        private void TranslteDialog_Load(object sender, EventArgs e)
        {
           
        }


        private async void BtnTranslate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(comboLanguages.Text))
            {
                MessageBox.Show("Please select a language.");
                return;
            }

            string prompt = $"Translate the following email into {comboLanguages.Text}. Return only the translated version:\n\n{originalText}";

            using (var loading = new LoadingDialog("Translating email..."))
            {
                loading.Show();
                Application.DoEvents();
                string result = await Ribbon2.CallGeminiAsyncStatic(prompt);
                loading.Close();



                if (result != "Error")
                {
                    txtPreview.Text = result;
                }
                else
                {
                    MessageBox.Show("Translation failed.");
                }
            }
        }

        private List<string> GetLanguageList()
        {
            return new List<string> {
            "English", "Malayalam", "Hindi", "Spanish", "French", "German", "Chinese", "Japanese", "Korean", "Italian",
            "Arabic", "Russian", "Portuguese", "Bengali", "Turkish", "Urdu", "Dutch", "Malay",
            "Vietnamese", "Swedish", "Polish", "Greek", "Tamil", "Telugu", "Gujarati"
        };
        }
    }
}
