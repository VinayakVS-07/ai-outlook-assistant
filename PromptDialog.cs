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
    public partial class PromptDialog : Form
    {
        public string UserInput { get; private set; }

        public PromptDialog(string title, string prompt)
        {
            InitializeComponent();

            this.Text = title;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Size = new Size(600, 300);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label lblPrompt = new Label()
            {
                Text = prompt,
                Left = 15,
                Top = 20,
                Width = 550,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            TextBox txtInput = new TextBox()
            {
                Multiline = true,
                Left = 15,
                Top = 50,
                Width = 550,
                Height = 140,
                Font = new Font("Segoe UI", 10),
                ScrollBars = ScrollBars.Vertical
            };

            Button btnOk = new Button()
            {
                Text = "OK",
                Left = 385,
                Width = 80,
                Top = 210,
                DialogResult = DialogResult.OK
            };

            Button btnCancel = new Button()
            {
                Text = "Cancel",
                Left = 475,
                Width = 80,
                Top = 210,
                DialogResult = DialogResult.Cancel
            };

            btnOk.Click += (sender, e) => { this.UserInput = txtInput.Text.Trim(); this.Close(); };
            btnCancel.Click += (sender, e) => { this.UserInput = null; this.Close(); };

            this.Controls.Add(lblPrompt);
            this.Controls.Add(txtInput);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
        }

        private void PromptDialog_Load(object sender, EventArgs e)
        {

        }
    }
}
