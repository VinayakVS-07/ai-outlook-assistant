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
    public partial class LoadingDialog : Form
    {
        public LoadingDialog(string message = "Processing... Please wait")
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ControlBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Width = 300;
            this.Height = 100;

            var label = new Label
            {
                Text = message,
                AutoSize = false,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };

            this.Controls.Add(label);
        }

        private void LoadingDialog_Load(object sender, EventArgs e)
        {

        }
    }
}


