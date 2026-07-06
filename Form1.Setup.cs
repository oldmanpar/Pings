using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Pings.Models;
using Pings.Utils;

namespace Pings
{
    public partial class Form1
    {
        // ====================================================================
        // UI構築
        // ====================================================================

        private void InitializeCustomComponents()
        {
            this.Text = "Pings";
            this.Size = new Size(1300, 700);
            this.StartPosition = FormStartPosition.CenterScreen;

            uiUpdateTimer = new System.Windows.Forms.Timer();
            uiUpdateTimer.Interval = UiUpdateInterval;
            uiUpdateTimer.Tick += UiUpdateTimer_Tick;

            // 上部パネル（日時・設定）
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(10) };
            Label lblStart = new Label { Text = "開始日時", Location = new Point(10, 5), AutoSize = true };
            txtStartTime = new TextBox { Location = new Point(10, 23), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };
            Label lblEnd = new Label { Text = "終了日時", Location = new Point(170, 5), AutoSize = true };
            txtEndTime = new TextBox { Location = new Point(170, 23), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };
            topPanel.Controls.AddRange(new Control[] { lblStart, txtStartTime, lblEnd, txtEndTime });

            Label lblInterval = new Label { Text = "送信間隔 [ms]", Location = new Point(350, 5), AutoSize = true };
            cmbInterval = new ComboBox { Location = new Point(350, 23), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbInterval.Items.AddRange(new object[] { "100", "500", "1000", "2000" });
            cmbInterval.SelectedIndex = 3; // 初期値 2000ms（多拠点監視を想定）

            Label lblTimeout = new Label { Text = "タイムアウト [ms]", Location = new Point(480, 5), AutoSize = true };
            cmbTimeout = new ComboBox { Location = new Point(480, 23), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTimeout.Items.AddRange(new object[] { "500", "1000", "2000", "5000" });
            cmbTimeout.SelectedIndex = 2;
            topPanel.Controls.AddRange(new Control[] { lblInterval, cmbInterval, lblTimeout, cmbTimeout });

            // インターフェイス選択（非表示）
            Label lblInterface = new Label { Text = "送信インターフェイス", Location = new Point(600, 5), AutoSize = true, Visible = false };
            cmbInterface = new ComboBox { Location = new Point(600, 23), Width = 240, DropDownStyle = ComboBoxStyle.DropDownList, Visible = false, Enabled = false };
            btnRefreshInterfaces = new Button { Text = "更新", Location = new Point(848, 22), Width = 50, Visible = false, Enabled = false };
            topPanel.Controls.AddRange(new Control[] { lblInterface, cmbInterface, btnRefreshInterfaces });

            // コンテンツパネル
            Panel contentPanel = new Panel { Dock = DockStyle.Fill };

            tabControl = new TabControl { Dock = DockStyle.Fill };

            // 監視統計タブ
            statsPage = new TabPage("監視統計");
            dgvMonitor = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = SystemColors.ControlLight,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToResizeRows = false,
                RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            };
            dgvMonitor.CellDoubleClick += DgvMonitor_CellDoubleClick;
            dgvMonitor.CellValueChanged += DgvMonitor_CellValueChanged;
            dgvMonitor.CurrentCellDirtyStateChanged += DgvMonitor_CurrentCellDirtyStateChanged;
            dgvMonitor.ColumnHeaderMouseDoubleClick += DgvMonitor_ColumnHeaderMouseDoubleClick;
            statsPage.Controls.Add(dgvMonitor);
            tabControl.Controls.Add(statsPage);

            // 障害ログタブ
            TabPage logPage = new TabPage("障害イベントログ");
            dgvLog = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                ReadOnly = true,
                BackgroundColor = SystemColors.ControlLight,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToDeleteRows = false,
                AllowUserToAddRows = false,
                AllowUserToResizeRows = false,
                RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            };
            logPage.Controls.Add(dgvLog);
            tabControl.Controls.Add(logPage);

            // Tracerouteタブ
            traceroutePage = new TabPage("Traceroute");
            traceroutePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                ColumnCount = 1,
                RowCount = 1
            };
            traceroutePage.Controls.Add(traceroutePanel);
            tabControl.Controls.Add(traceroutePage);

            // ステータスボードタブ（多拠点の一括確認用）
            tabControl.Controls.Add(CreateStatusBoardTab());

            contentPanel.Controls.Add(tabControl);
            contentPanel.Controls.Add(topPanel);

            // 下部パネル
            Panel bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 80, Margin = new Padding(0) };

            GroupBox gbPing = new GroupBox { Text = "Ping", Width = 460, Padding = new Padding(8) };
            btnPingStart = new Button { Text = "Ping開始",    Location = new Point(8,   22), Width = 100 };
            btnStop      = new Button { Text = "Ping停止",    Location = new Point(116, 22), Width = 100 };
            btnClear     = new Button { Text = "Ping結果クリア", Location = new Point(224, 22), Width = 110 };
            btnSave      = new Button { Text = "Ping結果保存",  Location = new Point(340, 22), Width = 110 };
            gbPing.Controls.AddRange(new Control[] { btnPingStart, btnStop, btnClear, btnSave });

            GroupBox gbTrace = new GroupBox { Text = "Traceroute", Width = 760, Padding = new Padding(8) };
            btnTraceroute      = new Button { Text = "Traceroute実行", Location = new Point(8,   22), Width = 120 };
            btnStopTraceroute  = new Button { Text = "Traceroute停止", Location = new Point(136, 22), Width = 120 };
            btnClearTraceroute = new Button { Text = "Trace結果クリア",  Location = new Point(264, 22), Width = 140 };
            btnSaveTraceroute  = new Button { Text = "Trace結果保存",   Location = new Point(412, 22), Width = 140 };
            gbTrace.Controls.AddRange(new Control[] { btnTraceroute, btnStopTraceroute, btnClearTraceroute, btnSaveTraceroute });

            var leftFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = gbPing.Width + gbTrace.Width + 24,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8),
                Margin = new Padding(0)
            };
            leftFlow.Controls.AddRange(new Control[] { gbPing, gbTrace });

            var rightPanel = new Panel { Dock = DockStyle.Right, Width = 120, Padding = new Padding(8) };
            btnExit = new Button { Text = "終了", Anchor = AnchorStyles.Top | AnchorStyles.Right, Width = 80, Height = 30 };
            btnExit.Location = new Point(rightPanel.ClientSize.Width - btnExit.Width - 10, 20);
            rightPanel.Resize += (s, e) =>
            {
                btnExit.Location = new Point(Math.Max(8, rightPanel.ClientSize.Width - btnExit.Width - 10), 20);
            };
            rightPanel.Controls.Add(btnExit);

            bottomPanel.Controls.AddRange(new Control[] { rightPanel, leftFlow });

            if (this.MainMenuStrip != null && !this.Controls.Contains(this.MainMenuStrip))
            {
                this.Controls.Add(this.MainMenuStrip);
                this.MainMenuStrip.Dock = DockStyle.Top;
            }
            this.Controls.Add(contentPanel);
            this.Controls.Add(bottomPanel);

            // イベント登録
            btnPingStart.Click       += BtnPingStart_Click;
            btnStop.Click            += BtnStop_Click;
            btnClear.Click           += BtnClear_Click;
            btnExit.Click            += BtnExit_Click;
            btnSave.Click            += BtnSaveResult_Click;
            btnTraceroute.Click      += BtnTraceroute_Click;
            btnSaveTraceroute.Click  += BtnSaveTraceroute_Click;
            btnClearTraceroute.Click += BtnClearTraceroute_Click;
            btnStopTraceroute.Click  += BtnStopTraceroute_Click;

            int menuHeight = this.MainMenuStrip?.PreferredSize.Height ?? 0;
            contentPanel.Padding = new Padding(0, menuHeight, 0, 0);

            tracerouteTextBoxes = new Dictionary<string, TextBox>(StringComparer.OrdinalIgnoreCase);
            UpdateTracerouteButtons();
        }

        // ====================================================================
        // メニューバー構築
        // ====================================================================

        private void SetupMenuStrip()
        {
            var ms = new MenuStrip { Dock = DockStyle.Top };

            var file = new ToolStripMenuItem("ファイル");
            file.DropDownItems.Add("対象アドレス保存", null, BtnSaveAddress_Click);
            file.DropDownItems.Add("対象アドレス読込", null, BtnLoadAddress_Click);
            file.DropDownItems.Add(new ToolStripSeparator());
            file.DropDownItems.Add("終了", null, BtnExit_Click);
            ms.Items.Add(file);

            var opt = new ToolStripMenuItem("オプション");

            mnuAutoSaveAllPing = new ToolStripMenuItem("Ping結果の自動保存") { CheckOnClick = true, Checked = true };
            mnuAutoSaveAllPing.CheckedChanged += (s, e) => UpdateUiState(CurrentUiState);
            opt.DropDownItems.Add(mnuAutoSaveAllPing);

            mnuAutoSaveTraceroute = new ToolStripMenuItem("Trace結果の自動保存") { CheckOnClick = true, Checked = true };
            mnuAutoSaveTraceroute.CheckedChanged += (s, e) => UpdateTracerouteButtons();
            opt.DropDownItems.Add(mnuAutoSaveTraceroute);

            opt.DropDownItems.Add(new ToolStripSeparator());

            mnuShowJitter1 = new ToolStripMenuItem("ジッタ①（最大値‐最小値）を表示させる") { CheckOnClick = true, Checked = true };
            mnuShowJitter1.CheckedChanged += (s, e) => UpdateColumnVisibility();
            opt.DropDownItems.Add(mnuShowJitter1);

            mnuShowJitter2 = new ToolStripMenuItem("ジッタ②（パケットペアの平均値）を表示させる") { CheckOnClick = true, Checked = false };
            mnuShowJitter2.CheckedChanged += (s, e) => UpdateColumnVisibility();
            opt.DropDownItems.Add(mnuShowJitter2);

            mnuShowStdDev = new ToolStripMenuItem("Pingの標準偏差を表示する") { CheckOnClick = true, Checked = false };
            mnuShowStdDev.CheckedChanged += (s, e) => UpdateColumnVisibility();
            opt.DropDownItems.Add(mnuShowStdDev);

            opt.DropDownItems.Add(new ToolStripSeparator());
            mnuTracerouteNoResolve = new ToolStripMenuItem("tracerouteで名前解決を行わない") { CheckOnClick = true, Checked = true };
            opt.DropDownItems.Add(mnuTracerouteNoResolve);

            ms.Items.Add(opt);

            var help = new ToolStripMenuItem("ヘルプ");
            help.DropDownItems.Add("バージョン情報", null, VersionItem_Click);
            ms.Items.Add(help);

            this.MainMenuStrip = ms;
            this.Controls.Add(ms);
        }

        // ====================================================================
        // DataGridView カラム定義
        // ====================================================================

        private void SetupDataGridViewColumns()
        {
            dgvMonitor.Columns.Clear();
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "ステータス",   HeaderText = "ｽﾃｰﾀｽ",    Width = 60,  ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = "Trace",       HeaderText = "Trace",     Width = 60,  TrueValue = true, FalseValue = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "順番",        HeaderText = "順番",      Width = 50,  ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "対象アドレス", HeaderText = "対象アドレス", Width = 120 });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "Host名",      HeaderText = "Host名",    Width = 120 });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "送信回数",    HeaderText = "送信回数",   Width = 80,  ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "失敗回数",    HeaderText = "失敗回数",   Width = 80,  ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "連続失敗回数", HeaderText = "連続失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "連続失敗時間s", HeaderText = "連続失敗時間", Width = 130, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "最大失敗時間s", HeaderText = "最大失敗時間", Width = 130, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "時間ms",      HeaderText = "時間[ms]",  Width = 80,  ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "平均値ms",    HeaderText = "平均値[ms]", Width = 80,  DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" }, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "最小値ms",    HeaderText = "最小値[ms]", Width = 80,  ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "最大値ms",    HeaderText = "最大値[ms]", Width = 80,  ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "Jitter1_MaxMin", Name = ColJitter1, HeaderText = "ジッタ①[ms]", Width = 80, ReadOnly = true, Visible = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "Jitter2_PktPair", Name = ColJitter2, HeaderText = "ジッタ②[ms]", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" }, ReadOnly = true, Visible = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn  { DataPropertyName = "StdDev",      Name = ColStdDev,  HeaderText = "標準偏差",    Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "F2" }, ReadOnly = true, Visible = false });

            dgvMonitor.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (e.ColumnIndex == 0 && dgvMonitor.Rows[e.RowIndex].DataBoundItem is PingMonitorItem item)
                {
                    var status = item.ステータス ?? "";
                    if (status.StartsWith("Down", StringComparison.OrdinalIgnoreCase))
                    {
                        e.CellStyle.BackColor = Color.MistyRose;
                        e.CellStyle.SelectionBackColor = Color.Red;
                    }
                    else if (status.StartsWith("復旧"))
                    {
                        e.CellStyle.BackColor = Color.LightSkyBlue;
                        e.CellStyle.SelectionBackColor = Color.Blue;
                    }
                    else if (status.Equals("OK", StringComparison.OrdinalIgnoreCase))
                    {
                        e.CellStyle.BackColor = Color.LightGreen;
                        e.CellStyle.SelectionBackColor = Color.Green;
                    }
                    else
                    {
                        e.CellStyle.BackColor = dgvMonitor.DefaultCellStyle.BackColor;
                        e.CellStyle.SelectionBackColor = dgvMonitor.DefaultCellStyle.SelectionBackColor;
                    }
                }
            };

            disruptionLogList = new SortableBindingList<DisruptionLogItem>(new System.Collections.Generic.List<DisruptionLogItem>());
            dgvLog.DataSource = disruptionLogList;
            dgvLog.ColumnHeaderMouseClick += dgvLog_ColumnHeaderMouseClick;
            dgvLog.Columns.Clear();
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "対象アドレス",  HeaderText = "対象アドレス",  Width = 120 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Host名",       HeaderText = "Host名",       Width = 120 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down開始日時",  HeaderText = "Down開始日時",  Width = 150 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧日時",      HeaderText = "復旧日時",      Width = 150 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗回数",      HeaderText = "失敗回数",      Width = 80  });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗時間mmss",  HeaderText = "失敗時間",      Width = 100 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前平均ms",  HeaderText = "Down前平均",    Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最小ms",  HeaderText = "Down前最小",    Width = 110 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最大ms",  HeaderText = "Down前最大",    Width = 110 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後平均ms",  HeaderText = "復旧後平均",    Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最小ms",  HeaderText = "復旧後最小",    Width = 110 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最大ms",  HeaderText = "復旧後最大",    Width = 110 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前Jitter1", Name = ColLogJitter1Pre,  HeaderText = "Down前ジッタ①",  Width = 100, Visible = false });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前Jitter2", Name = ColLogJitter2Pre,  HeaderText = "Down前ジッタ②",  Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" }, Visible = false });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前StdDev",  Name = ColLogStdDevPre,   HeaderText = "Down前標準偏差",  Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "F2" }, Visible = false });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後Jitter1", Name = ColLogJitter1Post, HeaderText = "復旧後ジッタ①",  Width = 100, Visible = false });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後Jitter2", Name = ColLogJitter2Post, HeaderText = "復旧後ジッタ②",  Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" }, Visible = false });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後StdDev",  Name = ColLogStdDevPost,  HeaderText = "復旧後標準偏差",  Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "F2" }, Visible = false });
        }

        // ====================================================================
        // カラム表示切り替え
        // ====================================================================

        private void UpdateColumnVisibility()
        {
            bool showJ1 = mnuShowJitter1.Checked;
            bool showJ2 = mnuShowJitter2.Checked;
            bool showSD = mnuShowStdDev.Checked;

            if (dgvMonitor.Columns.Contains(ColJitter1)) dgvMonitor.Columns[ColJitter1].Visible = showJ1;
            if (dgvMonitor.Columns.Contains(ColJitter2)) dgvMonitor.Columns[ColJitter2].Visible = showJ2;
            if (dgvMonitor.Columns.Contains(ColStdDev))  dgvMonitor.Columns[ColStdDev].Visible  = showSD;

            if (dgvLog.Columns.Contains(ColLogJitter1Pre)) dgvLog.Columns[ColLogJitter1Pre].Visible = showJ1;
            if (dgvLog.Columns.Contains(ColLogJitter2Pre)) dgvLog.Columns[ColLogJitter2Pre].Visible = showJ2;
            if (dgvLog.Columns.Contains(ColLogStdDevPre))  dgvLog.Columns[ColLogStdDevPre].Visible  = showSD;

            if (dgvLog.Columns.Contains(ColLogJitter1Post)) dgvLog.Columns[ColLogJitter1Post].Visible = showJ1;
            if (dgvLog.Columns.Contains(ColLogJitter2Post)) dgvLog.Columns[ColLogJitter2Post].Visible = showJ2;
            if (dgvLog.Columns.Contains(ColLogStdDevPost))  dgvLog.Columns[ColLogStdDevPost].Visible  = showSD;
        }
    }
}
