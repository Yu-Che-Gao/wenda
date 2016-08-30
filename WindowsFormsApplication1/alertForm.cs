using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class alertForm : Form
    {
        private string title = "";
        private string content = "";

        public alertForm()
        {
            InitializeComponent();
        }

        private void alertForm_Load(object sender, EventArgs e)
        {
            this.Text = title;
            label1.Text = content;
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        private void alertForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }

        public void setTitle(string title)
        {
            this.title = title;
        }

        public void setContent(string content)
        {
            this.content = content;
        }
    }
}
