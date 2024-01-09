using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ImportObjectsFromExcelUsingEPPlus
{
    public partial class StatusWindow : Form
    {

        public string Message { set { textProgressMessage.Text = value; } }

        Loggerton Logs = null;

        public StatusWindow(String caption)
        {            
            InitializeComponent();
            this.Text = caption;
            textLogs.Text = "";
            this.Refresh();
        }

        public void UpdateProgress(Int32 curProgress, string msg)
        {
            progressBar1.Value = curProgress;
            progressBar1.Refresh();
            textProgressMessage.Text = msg;
        }

        public void UpdateLogs(Loggerton logs)
        {
            textLogs.Text = logs.GetLogs(EnumLogFlags.All);
        }

        private void StatusWindow_Load(object sender, EventArgs e)
        {

        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
