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
    public partial class ReplyDialog : Form
    {
        public string OriginalEmail => txtOriginalEmail.Text.Trim();
        public string ReplyInstruction => txtReplyInstruction.Text.Trim();
        public string GeneratedReply { get; set; }
        public ReplyDialog()
        {
            InitializeComponent();
            this.Text = "AI Reply Assistant";
            this.Width = 600;
            this.Height = 450;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;

            var lblOriginal = new Label
            {
                Text = "📋 Email to Reply To:",
                Top = 10,
                Left = 10,
                Width = 560
            };
            txtOriginalEmail = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Left = 10,
                Top = 30,
                Width = 560,
                Height = 120
            };

            var lblInstruction = new Label
            {
                Text = "🧠 Reply Instructions (e.g., tone or intent):",
                Top = 160,
                Left = 10,
                Width = 560
            };
            txtReplyInstruction = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Left = 10,
                Top = 180,
                Width = 560,
                Height = 80
            };

            var btnGenerate = new Button
            {
                Text = "Generate Reply",
                Left = 10,
                Top = 270,
                Width = 150
            };
            btnGenerate.Click += BtnGenerate_Click;

            txtPreview = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Left = 10,
                Top = 310,
                Width = 560,
                Height = 60,
                ReadOnly = true,
                BackColor = System.Drawing.Color.White
            };

            var btnOK = new Button
            {
                Text = "Confirm",
                Left = 400,
                Top = 380,
                Width = 80,
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Left = 490,
                Top = 380,
                Width = 80,
                DialogResult = DialogResult.Cancel
            };

            this.Controls.Add(lblOriginal);
            this.Controls.Add(txtOriginalEmail);
            this.Controls.Add(lblInstruction);
            this.Controls.Add(txtReplyInstruction);
            this.Controls.Add(btnGenerate);
            this.Controls.Add(txtPreview);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);
        }

        private async void BtnGenerate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(OriginalEmail) || string.IsNullOrWhiteSpace(ReplyInstruction))
            {
                MessageBox.Show("Please fill in both the email and reply instruction.");
                return;
            }

            string prompt = $"Reply to the following email:\n{OriginalEmail}\n\nInstruction: {ReplyInstruction}\n\nReturn your response in this format:\nSubject: ...\nBody: ...";


            using (var loading = new LoadingDialog("Generating reply..."))
            {
                loading.Show();
                Application.DoEvents();
                GeneratedReply = await Ribbon2.CallGeminiAsyncStatic(prompt);
                loading.Close();
            }

            txtPreview.Text = GeneratedReply != "Error" ? GeneratedReply : "Reply generation failed.";
        }

        private TextBox txtOriginalEmail;
        private TextBox txtReplyInstruction;
        private TextBox txtPreview;

        private void ReplyDialog_Load(object sender, EventArgs e)
        {

        }
    }
}
