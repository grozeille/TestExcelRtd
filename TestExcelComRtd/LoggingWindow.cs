using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestExcelComRtd
{
    public partial class LoggingWindow : Form
    {
        private static LoggingWindow instance;

        public static void WriteLine(string format, params object[] args)
        {
            if (instance == null)
            {
                instance = new LoggingWindow();
                instance.Show();
            }
            instance.Show();

            instance.listBoxMessages.Items.Add(string.Format(format, args));
        }

        public LoggingWindow()
        {
            InitializeComponent();
        }

        private void LoggingWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }
    }
}
