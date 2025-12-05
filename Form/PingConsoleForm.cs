using System;
using System.Drawing;
using System.Windows.Forms;

namespace Pings
{
    /// <summary>
    /// 各ターゲット専用の簡易コンソール風ウィンドウ
    /// </summary>
    public class PingConsoleForm : Form
    {
        private TextBox txtConsole;
        public string Address { get; }

        public PingConsoleForm(string address, string hostName = "")
        {
            Address = address;
            InitializeComponent(address, hostName);
        }

        private void InitializeComponent(string address, string hostName)
        {
            this.Text = $"Ping Console - {address}{(string.IsNullOrEmpty(hostName) ? "" : " (" + hostName + ")")}";
            this.Size = new Size(520, 240);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;

            txtConsole = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = false,
                Font = new Font(FontFamily.GenericMonospace, 9f),
                BackColor = SystemColors.Window
            };

            this.Controls.Add(txtConsole);
        }

        /// <summary>
        /// スレッドセーフに行を追記して末尾にスクロールする
        /// </summary>
        public void AppendLine(string line)
        {
            if (this.IsDisposed) return;
            if (txtConsole.InvokeRequired)
            {
                txtConsole.Invoke(new Action(() => AppendLineInternal(line)));
            }
            else
            {
                AppendLineInternal(line);
            }
        }

        private void AppendLineInternal(string line)
        {
            if (this.IsDisposed) return;
            txtConsole.AppendText(line + Environment.NewLine);
            txtConsole.SelectionStart = txtConsole.Text.Length;
            txtConsole.ScrollToCaret();
        }
    }
}