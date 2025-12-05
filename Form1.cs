using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Pings.Models;
using Pings.Repositories;
using Pings.Services;
using Pings.Utils;

namespace Pings
{
    public partial class Form1 : Form
    {
        // -------------------------
        // 依存オブジェクト (Service / Repository)
        // -------------------------
        private readonly PingFileRepository _repository;
        private readonly PingService _pingService;
        private readonly TracerouteService _tracerouteService;

        // -------------------------
        // データバインディング
        // -------------------------
        private SortableBindingList<DisruptionLogItem> disruptionLogList;
        private BindingList<PingMonitorItem> monitorList;

        // -------------------------
        // UIコントロール (デザイナーコードの代わり)
        // -------------------------
        private DataGridView dgvMonitor;
        private DataGridView dgvLog;
        private TextBox txtStartTime, txtEndTime;
        private ComboBox cmbInterval, cmbTimeout;
        private Button btnPingStart, btnStop, btnClear, btnSave, btnExit;
        private TabControl tabControl;
        private TabPage traceroutePage;
        private TableLayoutPanel traceroutePanel;
        private Button btnTraceroute, btnSaveTraceroute, btnClearTraceroute, btnStopTraceroute;
        private ToolStripMenuItem mnuAutoSaveAllPing, mnuAutoSaveTraceroute;

        // インターフェイス選択用 (非表示機能)
        private ComboBox cmbInterface;
        private Button btnRefreshInterfaces;
        private IPAddress _selectedLocalAddress;

        // -------------------------
        // 状態管理フィールド
        // -------------------------
        private CancellationTokenSource cts;
        private CancellationTokenSource tracerouteCts;
        private System.Windows.Forms.Timer uiUpdateTimer;
        private int _nextIndex = 1;
        private bool _allowEditAfterStop = true;
        private bool _allowIntervalTimeoutEdit = true;
        private bool _autoSaveFailed = false;
        private volatile bool _isTracerouteRunning = false;

        // Traceroute関連
        private Dictionary<string, TextBox> tracerouteTextBoxes;
        private SemaphoreSlim tracerouteSemaphore = new SemaphoreSlim(4);
        private const int TracerouteColumnWidth = 480;
        private ConcurrentDictionary<string, bool> _tracerouteCompletion;
        private ConcurrentDictionary<string, bool> _tracerouteStoppedByUser;

        // ソート状態
        private DataGridViewColumn currentSortColumn = null;
        private ListSortDirection currentSortDirection = ListSortDirection.Ascending;

        private const int UiUpdateInterval = 1000; // 例: 1000ms（1秒ごとにUI更新）

        public Form1()
        {
            // Form1 コンストラクタ（順序を入れ替える）
            InitializeComponent();
            SetupMenuStrip();              // ← 先にメニューを作る
            InitializeCustomComponents();
            SetupDataGridViewColumns();

            // 依存オブジェクトの生成
            _repository = new PingFileRepository();
            _tracerouteService = new TracerouteService();
            // PingServiceには「ログが作成されたときにUIへ追加する処理」を渡す
            _pingService = new PingService((logItem) =>
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => AddDisruptionLogItem(logItem)));
                }
                else
                {
                    AddDisruptionLogItem(logItem);
                }
            });

            // リスト初期化
            monitorList = new BindingList<PingMonitorItem>();
            monitorList.ListChanged += MonitorList_ListChanged;
            dgvMonitor.DataSource = monitorList;

            dgvMonitor.UserDeletedRow += DgvMonitor_UserDeletedRow;
            dgvMonitor.RowValidated += DgvMonitor_RowValidated;
            dgvMonitor.DefaultValuesNeeded += DgvMonitor_DefaultValuesNeeded;
            this.FormClosing += Form1_FormClosing;

            // 初期データ
            monitorList.Add(new PingMonitorItem(_nextIndex++, "127.0.0.1", "loopback"));
            monitorList.Add(new PingMonitorItem(_nextIndex++, "8.8.8.8", "Google DNS1"));
            monitorList.Add(new PingMonitorItem(_nextIndex++, "8.8.4.4", "Google DNS2"));

            UpdateUiState("Initial");
        }

        // ====================================================================
        // イベントハンドラ (UI → Logic呼び出し)
        // ====================================================================

        private void BtnPingStart_Click(object sender, EventArgs e)
        {
            StopMonitoring();
            ClearView();
            dgvMonitor.EndEdit();

            // 空行削除
            var validItems = monitorList
                .Where(i => !string.IsNullOrWhiteSpace(i.対象アドレス) || !string.IsNullOrWhiteSpace(i.Host名))
                .ToList();

            monitorList.ListChanged -= MonitorList_ListChanged;
            monitorList.Clear();
            foreach (var item in validItems) monitorList.Add(item);
            monitorList.ListChanged += MonitorList_ListChanged;

            _nextIndex = monitorList.Any() ? monitorList.Max(i => i.順番) + 1 : 1;

            StartMonitoring();
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            StopMonitoring();
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            StopMonitoring();
            ClearView();
            _allowEditAfterStop = true;
            _allowIntervalTimeoutEdit = true;
            UpdateUiState("Initial");
            btnPingStart.Enabled = true;
            dgvMonitor.AllowUserToDeleteRows = true;
        }

        private void BtnSaveResult_Click(object sender, EventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中は結果を保存できません。", "保存不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 既定パスでの保存を試みる
            string baseFolder = AppDomain.CurrentDomain.BaseDirectory;
            string defaultPath = Path.Combine(baseFolder, "Ping_Result", $"Ping_Result_{DateTime.Now:yyyyMMdd}.csv");

            try
            {
                ExecuteSave(defaultPath);
                MessageBox.Show($"監視結果を保存しました：\n{defaultPath}", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch
            {
                // 既定失敗時は手動保存
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSVファイル (*.csv)|*.csv|All|*.*";
                sfd.FileName = Path.GetFileName(defaultPath);
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExecuteSave(sfd.FileName);
                        MessageBox.Show($"監視結果を保存しました：\n{sfd.FileName}", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"保存エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            void ExecuteSave(string path)
            {
                _repository.SavePingResults(path, monitorList, disruptionLogList, txtStartTime.Text, txtEndTime.Text, cmbInterval.Text, cmbTimeout.Text);
                _autoSaveFailed = false;
                _allowEditAfterStop = true;
                _allowIntervalTimeoutEdit = true;
                UpdateUiState("Stopped");
                btnPingStart.Enabled = true;
            }
        }

        private void BtnSaveAddress_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSVファイル (*.csv)|*.csv|All|*.*";
                sfd.FileName = "PingTargets.csv";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _repository.SaveAddresses(sfd.FileName, monitorList);
                        MessageBox.Show("監視対象アドレスを保存しました。", "保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"保存エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnLoadAddress_Click(object sender, EventArgs e)
        {
            if (cts != null) return;

            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "アドレスファイル (*.csv;*.txt;*.log)|*.csv;*.txt;*.log|All|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var items = _repository.LoadAddresses(ofd.FileName);
                        monitorList = new BindingList<PingMonitorItem>(items);
                        monitorList.ListChanged += MonitorList_ListChanged;
                        dgvMonitor.DataSource = monitorList;
                        _nextIndex = items.Count + 1;
                        UpdateUiState("Initial");
                        MessageBox.Show($"監視対象アドレスを {items.Count}件 読み込みました。", "読込完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"読込エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ====================================================================
        // Traceroute 関連ハンドラ
        // ====================================================================

        private async void BtnTraceroute_Click(object sender, EventArgs e)
        {
            var targets = monitorList.Where(i => i.Trace && !string.IsNullOrWhiteSpace(i.対象アドレス))
                                     .Select(i => i.対象アドレス).Distinct().ToList();

            if (!targets.Any())
            {
                MessageBox.Show("Traceroute対象がありません。", "実行不可", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _isTracerouteRunning = true;
            UpdateUiState(cts != null ? "Running" : (string.IsNullOrEmpty(txtStartTime.Text) ? "Initial" : "Stopped"));
            UpdateTracerouteButtons();
            tabControl.SelectedTab = traceroutePage;

            // UI構築
            SetupTracerouteUI(targets);

            AppendTracerouteOutputToAll("================================================================\r\n", false);
            AppendTracerouteOutputToAll($"== Traceroute: {DateTime.Now:yyyy/MM/dd HH:mm:ss} ==\r\n", false);

            tracerouteCts?.Dispose();
            tracerouteCts = CancellationTokenSource.CreateLinkedTokenSource(cts?.Token ?? CancellationToken.None);

            _tracerouteCompletion = new ConcurrentDictionary<string, bool>();
            _tracerouteStoppedByUser = new ConcurrentDictionary<string, bool>();
            foreach (var t in targets)
            {
                _tracerouteCompletion[t] = false;
                _tracerouteStoppedByUser[t] = false;
            }

            try
            {
                int timeout = 4000;
                int.TryParse(cmbTimeout.Text, out timeout);

                var tasks = targets.Select(async address =>
                {
                    await tracerouteSemaphore.WaitAsync(tracerouteCts.Token);
                    try
                    {
                        // サービスの呼び出し (コールバックでUI更新)
                        await _tracerouteService.RunTracerouteAsync(address, timeout, tracerouteCts.Token, (text) =>
                        {
                            AppendTracerouteOutput(address, text, false);
                        });
                    }
                    finally
                    {
                        tracerouteSemaphore.Release();
                        if (_tracerouteCompletion != null) _tracerouteCompletion[address] = true;

                        // 完了時のメッセージ
                        bool stopped = _tracerouteStoppedByUser.ContainsKey(address) && _tracerouteStoppedByUser[address];
                        if (!stopped && !tracerouteCts.IsCancellationRequested)
                        {
                            AppendTracerouteOutput(address, "=== Traceroute 完了 ===\r\n\r\n", false);
                        }
                    }
                });

                await Task.WhenAll(tasks);
            }
            catch (OperationCanceledException)
            {
                HandleTracerouteCancel(targets);
            }
            finally
            {
                if (tracerouteCts != null && tracerouteCts.IsCancellationRequested) HandleTracerouteCancel(targets);

                AppendTracerouteOutputToAll("================================================================\r\n\r\n", false);
                _isTracerouteRunning = false;
                UpdateUiState(cts != null ? "Running" : (string.IsNullOrEmpty(txtStartTime.Text) ? "Initial" : "Stopped"));
                UpdateTracerouteButtons();

                if (mnuAutoSaveTraceroute.Checked) SaveTracerouteOutputsAutoAppend();
                tracerouteCts?.Dispose();
                tracerouteCts = null;
            }
        }

        private void SetupTracerouteUI(List<string> targets)
        {
            // UI再利用判定
            bool reuse = tracerouteTextBoxes != null && tracerouteTextBoxes.Count == targets.Count
                         && targets.All(t => tracerouteTextBoxes.ContainsKey(t));

            if (!reuse)
            {
                tracerouteTextBoxes.Clear();
                traceroutePanel.Controls.Clear();
                traceroutePanel.ColumnCount = Math.Max(1, targets.Count);
                traceroutePanel.RowCount = 1;
                traceroutePanel.ColumnStyles.Clear();
                for (int i = 0; i < traceroutePanel.ColumnCount; i++)
                    traceroutePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, TracerouteColumnWidth));

                for (int i = 0; i < targets.Count; i++)
                {
                    string address = targets[i];
                    string hostName = monitorList.FirstOrDefault(m => m.対象アドレス == address)?.Host名 ?? "";

                    var container = new Panel { Dock = DockStyle.Fill, Padding = new Padding(2), Width = TracerouteColumnWidth };
                    var lbl = new Label { Text = $"{address}_{hostName}", Dock = DockStyle.Top, Height = 18, AutoEllipsis = true };
                    var tb = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill, BackColor = SystemColors.Window, Font = new Font(FontFamily.GenericMonospace, 9f) };

                    container.Controls.Add(tb);
                    container.Controls.Add(lbl);
                    traceroutePanel.Controls.Add(container, i, 0);
                    tracerouteTextBoxes[address] = tb;
                }
            }
        }

        private void HandleTracerouteCancel(List<string> targets)
        {
            foreach (var addr in targets)
            {
                bool done = _tracerouteCompletion != null && _tracerouteCompletion.TryGetValue(addr, out bool d) && d;
                bool stopped = _tracerouteStoppedByUser != null && _tracerouteStoppedByUser.TryGetValue(addr, out bool s) && s;
                if (!done && !stopped)
                {
                    AppendTracerouteOutput(addr, "=== Tracerouteは途中で停止されました。 ===\r\n\r\n", false);
                    if (_tracerouteStoppedByUser != null) _tracerouteStoppedByUser[addr] = true;
                }
            }
        }

        private void BtnStopTraceroute_Click(object sender, EventArgs e)
        {
            if (_tracerouteCompletion != null)
            {
                foreach (var kv in _tracerouteCompletion)
                {
                    if (!kv.Value && !_tracerouteStoppedByUser[kv.Key])
                    {
                        AppendTracerouteOutput(kv.Key, "=== Tracerouteは途中で停止されました。 ===\r\n\r\n", false);
                        _tracerouteStoppedByUser[kv.Key] = true;
                    }
                }
            }
            tracerouteCts?.Cancel();
            _isTracerouteRunning = false;
            UpdateTracerouteButtons();
        }

        private void BtnSaveTraceroute_Click(object sender, EventArgs e)
        {
            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Traceroute_Result");
            try
            {
                foreach (var kv in tracerouteTextBoxes)
                {
                    string host = monitorList.FirstOrDefault(i => i.対象アドレス == kv.Key)?.Host名 ?? "";
                    _repository.SaveTracerouteResult(folder, kv.Key, host, kv.Value.Text);
                    kv.Value.Clear();
                }
                MessageBox.Show("保存しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"エラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            UpdateTracerouteButtons();
        }

        private void SaveTracerouteOutputsAutoAppend()
        {
            string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Traceroute_Result");
            foreach (var kv in tracerouteTextBoxes)
            {
                string host = monitorList.FirstOrDefault(i => i.対象アドレス == kv.Key)?.Host名 ?? "";
                _repository.SaveTracerouteResult(folder, kv.Key, host, kv.Value.Text);
            }
        }

        private void BtnClearTraceroute_Click(object sender, EventArgs e)
        {
            if (tracerouteTextBoxes != null)
            {
                foreach (var tb in tracerouteTextBoxes.Values) tb.Clear();
                UpdateTracerouteButtons();
            }
        }

        private void AppendTracerouteOutput(string address, string text, bool keepTimestamp)
        {
            if (tracerouteTextBoxes != null && tracerouteTextBoxes.ContainsKey(address))
            {
                var tb = tracerouteTextBoxes[address];
                if (tb.InvokeRequired) tb.Invoke(new Action(() => { tb.AppendText(text); tb.ScrollToCaret(); }));
                else { tb.AppendText(text); tb.ScrollToCaret(); }
                UpdateTracerouteButtons();
            }
        }

        private void AppendTracerouteOutputToAll(string text, bool keepTimestamp)
        {
            if (tracerouteTextBoxes != null)
            {
                foreach (var addr in tracerouteTextBoxes.Keys) AppendTracerouteOutput(addr, text, keepTimestamp);
            }
        }

        // ====================================================================
        // 内部ロジック (Start/Stop, UpdateUI)
        // ====================================================================

        private void StartMonitoring()
        {
            _allowEditAfterStop = false;
            _allowIntervalTimeoutEdit = false;
            cts = new CancellationTokenSource();
            txtStartTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            txtEndTime.Text = "";

            int interval = int.Parse(cmbInterval.Text);
            int timeout = int.Parse(cmbTimeout.Text);

            UpdateUiState("Running");
            uiUpdateTimer.Start();
            ResetLogSortIndicators();

            foreach (var item in monitorList.Where(i => !string.IsNullOrEmpty(i.対象アドレス) && i.順番 > 0))
            {
                item.送信間隔ms = interval;
                item.タイムアウトms = timeout;
                item.ResetData();
                // サービスにPingループの実行を委譲
                Task.Run(() => _pingService.RunPingLoopAsync(item, cts.Token));
            }
            dgvMonitor.AllowUserToDeleteRows = false;
        }

        private void StopMonitoring()
        {
            if (cts != null)
            {
                uiUpdateTimer.Stop();
                cts.Cancel();
                cts.Dispose();
                cts = null;
                txtEndTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                _allowEditAfterStop = false;

                if (mnuAutoSaveAllPing.Checked)
                {
                    try
                    {
                        string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ping_Result", $"Ping_Result_{DateTime.Now:yyyyMMdd}.csv");
                        _repository.SavePingResults(path, monitorList, disruptionLogList, txtStartTime.Text, txtEndTime.Text, cmbInterval.Text, cmbTimeout.Text);
                        MessageBox.Show($"自動保存しました：\n{path}", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        _autoSaveFailed = true;
                        MessageBox.Show($"自動保存失敗: {ex.Message}\n手動保存してください。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                UpdateUiState("Stopped");
                btnPingStart.Enabled = false;
            }
        }

        private void ClearView()
        {
            foreach (var item in monitorList) item.ResetData();
            disruptionLogList.Clear();
            txtStartTime.Text = "";
            txtEndTime.Text = "";
            monitorList.ResetBindings();
            UpdateUiState("Initial");
            ResetLogSortIndicators();
            if (tracerouteTextBoxes != null)
            {
                tracerouteTextBoxes.Clear();
                traceroutePanel.Controls.Clear();
            }
        }

        // ====================================================================
        // UI状態更新・ヘルパー
        // ====================================================================

        private void UpdateUiState(string state)
        {
            // メニュー制御
            if (MainMenuStrip != null && MainMenuStrip.Items.Count > 0)
            {
                var fileMenu = MainMenuStrip.Items[0] as ToolStripMenuItem;
                if (fileMenu != null)
                {
                    fileMenu.DropDownItems[0].Enabled = (state == "Initial" || state == "Stopped");
                    fileMenu.DropDownItems[1].Enabled = (state == "Initial");
                }
            }

            dgvMonitor.ReadOnly = false;
            dgvMonitor.AllowUserToAddRows = (state == "Initial");

            foreach (DataGridViewColumn col in dgvMonitor.Columns)
            {
                if (state == "Initial")
                {
                    if (col.DataPropertyName == "対象アドレス" || col.DataPropertyName == "Host名" || col.DataPropertyName == "Trace") col.ReadOnly = false;
                    else col.ReadOnly = true;
                }
                else if (state == "Running")
                {
                    if (col.DataPropertyName == "Trace") col.ReadOnly = false;
                    else col.ReadOnly = true;
                }
                else // Stopped
                {
                    if (col.DataPropertyName == "Trace") col.ReadOnly = false;
                    else if (col.DataPropertyName == "対象アドレス" || col.DataPropertyName == "Host名") col.ReadOnly = !_allowEditAfterStop;
                    else col.ReadOnly = true;
                }
            }

            if (state != "Running") btnTraceroute.Enabled = false;
            UpdateTracerouteButtons();

            bool allowCombos = !_isTracerouteRunning && state != "Running" && _allowIntervalTimeoutEdit;
            cmbInterval.Enabled = allowCombos;
            cmbTimeout.Enabled = allowCombos;

            // null 安全にする
            bool autoSave = mnuAutoSaveAllPing?.Checked ?? false;

            switch (state)
            {
                case "Initial":
                    btnPingStart.Enabled = true;
                    btnStop.Enabled = false;
                    btnClear.Enabled = false;
                    btnSave.Enabled = false;
                    break;
                case "Running":
                    btnPingStart.Enabled = false;
                    btnStop.Enabled = true;
                    btnClear.Enabled = false;
                    btnSave.Enabled = false;
                    break;
                case "Stopped":
                    btnPingStart.Enabled = true;
                    btnStop.Enabled = false;
                    btnClear.Enabled = true;
                    btnSave.Enabled = !autoSave || _autoSaveFailed;
                    break;
            }

            if (_isTracerouteRunning)
            {
                btnPingStart.Enabled = false;
                btnStop.Enabled = false;
            }
        }

        // C#
        private void UpdateTracerouteButtons()
        {
            if (btnTraceroute == null) return;
            if (InvokeRequired) { Invoke((MethodInvoker)UpdateTracerouteButtons); return; }

            bool hasChecked = (monitorList?.Any(i => i.Trace && !string.IsNullOrWhiteSpace(i.対象アドレス)) ?? false);
            btnTraceroute.Enabled = hasChecked && !_isTracerouteRunning;

            bool hasOutput = (tracerouteTextBoxes != null) && tracerouteTextBoxes.Values.Any(t => !string.IsNullOrEmpty(t.Text));
            bool autoSave = mnuAutoSaveTraceroute?.Checked ?? false;

            btnSaveTraceroute.Enabled = !_isTracerouteRunning && hasOutput && !autoSave;
            btnClearTraceroute.Enabled = !_isTracerouteRunning && hasOutput;
            btnStopTraceroute.Enabled = _isTracerouteRunning;
        }
        private void AddDisruptionLogItem(DisruptionLogItem item)
        {
            disruptionLogList.Add(item);
            if (currentSortColumn == null || (currentSortColumn.DataPropertyName == "復旧日時" && currentSortDirection == ListSortDirection.Ascending))
            {
                if (dgvLog.RowCount > 0) dgvLog.FirstDisplayedScrollingRowIndex = dgvLog.RowCount - 1;
            }
        }

        private void UiUpdateTimer_Tick(object sender, EventArgs e)
        {
            if (dgvMonitor != null && dgvMonitor.IsHandleCreated) dgvMonitor.Refresh();
            if (dgvLog != null && dgvLog.IsHandleCreated) dgvLog.Refresh();
        }

        private bool _recalculating = false;

        private void RecalculateOrderNumbers()
        {
            if (_recalculating) return;
            _recalculating = true;
            try
            {
                monitorList.ListChanged -= MonitorList_ListChanged;
                for (int i = 0; i < monitorList.Count; i++) monitorList[i].順番 = i + 1;
                _nextIndex = monitorList.Count + 1;
                monitorList.ResetBindings();
                monitorList.ListChanged += MonitorList_ListChanged;
            }
            finally
            {
                _recalculating = false;
            }
        }

        // DataGridViewイベント
        private void MonitorList_ListChanged(object sender, ListChangedEventArgs e)
        {
            if (e.ListChangedType == ListChangedType.ItemAdded || e.ListChangedType == ListChangedType.ItemDeleted || e.ListChangedType == ListChangedType.Reset)
                RecalculateOrderNumbers();
        }
        private void DgvMonitor_UserDeletedRow(object sender, DataGridViewRowEventArgs e) => RecalculateOrderNumbers();
        private void DgvMonitor_RowValidated(object sender, DataGridViewCellEventArgs e) => RecalculateOrderNumbers();
        private void DgvMonitor_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e) { if (e.Row.IsNewRow) e.Row.Cells[2].Value = _nextIndex++; }
        private void DgvMonitor_CellValueChanged(object sender, DataGridViewCellEventArgs e) { if (e.ColumnIndex >= 0 && dgvMonitor.Columns[e.ColumnIndex].DataPropertyName == "Trace") UpdateTracerouteButtons(); }
        private void DgvMonitor_CurrentCellDirtyStateChanged(object sender, EventArgs e) { if (dgvMonitor.IsCurrentCellDirty) dgvMonitor.CommitEdit(DataGridViewDataErrorContexts.Commit); }
        private void DgvMonitor_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex >= 0 && dgvMonitor.Columns[e.ColumnIndex].DataPropertyName == "Trace")
            {
                bool val = monitorList.Any(i => !i.Trace);
                foreach (var item in monitorList) item.Trace = val;
                monitorList.ResetBindings();
                UpdateTracerouteButtons();
            }
        }
        // ★★★【追加】ダブルクリック時の処理 ★★★
        private void DgvMonitor_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // ヘッダーや無効な行のクリックは無視
            if (e.RowIndex < 0) return;

            // データバインドされているアイテムを取得 (安全なキャスト)
            if (dgvMonitor.Rows[e.RowIndex].DataBoundItem is PingMonitorItem item)
            {
                // アドレスが空の場合は何もしない
                if (string.IsNullOrWhiteSpace(item.対象アドレス)) return;

                // 新しいPingウィンドウを作成して表示
                // Ownerをthisにすると、メイン画面の前面に表示されやすくなります
                var terminal = new PingTerminalForm(item.対象アドレス, item.Host名);
                terminal.Show(this);
            }
        }
        private void dgvLog_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var col = dgvLog.Columns[e.ColumnIndex];
            var dir = (currentSortColumn == col && currentSortDirection == ListSortDirection.Ascending) ? ListSortDirection.Descending : ListSortDirection.Ascending;
            var prop = TypeDescriptor.GetProperties(typeof(DisruptionLogItem)).Find(col.DataPropertyName, true);
            if (prop != null)
            {
                if (currentSortColumn != null) currentSortColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                disruptionLogList.Sort(prop, dir);
                col.HeaderCell.SortGlyphDirection = (dir == ListSortDirection.Ascending) ? SortOrder.Ascending : SortOrder.Descending;
                currentSortColumn = col;
                currentSortDirection = dir;
            }
        }
        private void ResetLogSortIndicators()
        {
            if (currentSortColumn != null)
            {
                currentSortColumn.HeaderCell.SortGlyphDirection = SortOrder.None;
                currentSortColumn = null;
                currentSortDirection = ListSortDirection.Ascending;
            }
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中は終了できません。", "終了不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            StopMonitoring();
            Close();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (cts != null)
            {
                MessageBox.Show("Ping監視中は終了できません。", "終了不可", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
            else StopMonitoring();
        }
        private void VersionItem_Click(object sender, EventArgs e)
        {
            using (var ab = new AboutBox()) ab.ShowDialog(this);
        }

        // インターフェイス関連 (非表示機能だがコードは維持)
        private void PopulateNetworkInterfaces()
        {
            try
            {
                cmbInterface.Items.Clear();
                foreach (var ni in NetworkInterface.GetAllNetworkInterfaces().Where(n => n.NetworkInterfaceType != NetworkInterfaceType.Loopback && n.OperationalStatus == OperationalStatus.Up))
                {
                    foreach (var ip in ni.GetIPProperties().UnicastAddresses.Where(u => !IPAddress.IsLoopback(u.Address)))
                    {
                        cmbInterface.Items.Add(new { Display = $"{ni.Name} ({ip.Address})", Address = ip.Address });
                    }
                }
            }
            catch { }
        }

        // ====================================================================
        // UI初期化コード (長いので最後に記述)
        // ====================================================================
        private void InitializeCustomComponents()
        {
            this.Text = "Pings";
            this.Size = new Size(1300, 700);
            this.StartPosition = FormStartPosition.CenterScreen;

            uiUpdateTimer = new System.Windows.Forms.Timer();
            uiUpdateTimer.Interval = UiUpdateInterval;
            uiUpdateTimer.Tick += UiUpdateTimer_Tick;

            // Top panel (日時、設定)
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(10) };
            Label lblStart = new Label { Text = "開始日時", Location = new Point(10, 5), AutoSize = true };
            txtStartTime = new TextBox { Location = new Point(10, 23), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };
            Label lblEnd = new Label { Text = "終了日時", Location = new Point(170, 5), AutoSize = true };
            txtEndTime = new TextBox { Location = new Point(170, 23), Width = 150, ReadOnly = true, BackColor = SystemColors.ControlLight };
            topPanel.Controls.Add(lblStart);
            topPanel.Controls.Add(txtStartTime);
            topPanel.Controls.Add(lblEnd);
            topPanel.Controls.Add(txtEndTime);

            Label lblInterval = new Label { Text = "送信間隔 [ms]", Location = new Point(350, 5), AutoSize = true };
            cmbInterval = new ComboBox { Location = new Point(350, 23), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbInterval.Items.AddRange(new object[] { "100", "500", "1000", "2000" });
            cmbInterval.SelectedIndex = 2;

            Label lblTimeout = new Label { Text = "タイムアウト [ms]", Location = new Point(480, 5), AutoSize = true };
            cmbTimeout = new ComboBox { Location = new Point(480, 23), Width = 60, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTimeout.Items.AddRange(new object[] { "500", "1000", "2000", "5000" });
            cmbTimeout.SelectedIndex = 2;

            topPanel.Controls.Add(lblInterval);
            topPanel.Controls.Add(cmbInterval);
            topPanel.Controls.Add(lblTimeout);
            topPanel.Controls.Add(cmbTimeout);

            // 置換: InitializeCustomComponents 内のインターフェイス選択部（lblInterface / cmbInterface / btnRefreshInterfaces と PopulateNetworkInterfaces 呼び出し）を以下に差し替えてください。

            Label lblInterface = new Label { Text = "送信インターフェイス", Location = new Point(600, 5), AutoSize = true };
            cmbInterface = new ComboBox { Location = new Point(600, 23), Width = 240, DropDownStyle = ComboBoxStyle.DropDownList };
            btnRefreshInterfaces = new Button { Text = "更新", Location = new Point(848, 22), Width = 50 };

            // topPanel に追加
            topPanel.Controls.Add(lblInterface);
            topPanel.Controls.Add(cmbInterface);
            topPanel.Controls.Add(btnRefreshInterfaces);

            // 非表示にする（要求: コンボボックスを隠す）
            lblInterface.Visible = false;
            cmbInterface.Visible = false;
            btnRefreshInterfaces.Visible = false;

            // また念のため無効化してフォーカス対象外にする
            cmbInterface.Enabled = false;
            btnRefreshInterfaces.Enabled = false;

            // 初期一覧読み込みは行わず、必要な場合に手動で PopulateNetworkInterfaces を呼ぶようにする。
            //PopulateNetworkInterfaces();

            // content panel
            Panel contentPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0) };

            // tabControl inside content
            tabControl = new TabControl { Dock = DockStyle.Fill };
            TabPage statsPage = new TabPage("監視統計");
            dgvMonitor = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = SystemColors.ControlLight,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            // 変更: 行の高さをユーザーが変更できないようにする
            dgvMonitor.AllowUserToResizeRows = false;
            dgvMonitor.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

            // ★★★【追加】ここから ★★★
            // ダブルクリックイベントを登録
            dgvMonitor.CellDoubleClick += DgvMonitor_CellDoubleClick;
            // ★★★【追加】ここまで ★★★

            statsPage.Controls.Add(dgvMonitor);
            // DataGridView のチェック変更で Traceroute ボタンを即時更新するためのイベント登録
            dgvMonitor.CellValueChanged += DgvMonitor_CellValueChanged;
            dgvMonitor.CurrentCellDirtyStateChanged += DgvMonitor_CurrentCellDirtyStateChanged;
            dgvMonitor.ColumnHeaderMouseDoubleClick += DgvMonitor_ColumnHeaderMouseDoubleClick;
            tabControl.Controls.Add(statsPage);

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
                AllowUserToAddRows = false
            };
            // 変更: 行の高さをユーザーが変更できないようにする
            dgvLog.AllowUserToResizeRows = false;
            dgvLog.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

            logPage.Controls.Add(dgvLog);
            tabControl.Controls.Add(logPage);

            traceroutePage = new TabPage("Traceroute");

            // Traceroute 用の TableLayoutPanel を用意し、実行時に列を追加していく設計
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

            contentPanel.Controls.Add(tabControl);
            contentPanel.Controls.Add(topPanel);

            // bottom panel - グループで分けて配置
            Panel bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 80, Margin = new Padding(0) };

            // Ping 関連グループ
            GroupBox gbPing = new GroupBox { Text = "Ping", Dock = DockStyle.None, Width = 460, Padding = new Padding(8), AutoSize = false };
            btnPingStart = new Button { Text = "Ping開始", Location = new Point(8, 22), Width = 100 };
            btnStop = new Button { Text = "Ping停止", Location = new Point(116, 22), Width = 100 };
            btnClear = new Button { Text = "Ping結果クリア", Location = new Point(224, 22), Width = 110 }; // 名前変更
            btnSave = new Button { Text = "Ping結果保存", Location = new Point(340, 22), Width = 110 };

            gbPing.Controls.Add(btnPingStart);
            gbPing.Controls.Add(btnStop);
            gbPing.Controls.Add(btnClear);
            gbPing.Controls.Add(btnSave);

            // Traceroute 関連グループ
            GroupBox gbTrace = new GroupBox { Text = "Traceroute", Dock = DockStyle.None, Width = 760, Padding = new Padding(8), AutoSize = false };
            btnTraceroute = new Button { Text = "Traceroute実行", Location = new Point(8, 22), Width = 120 };
            btnStopTraceroute = new Button { Text = "Traceroute停止", Location = new Point(136, 22), Width = 120 };
            btnClearTraceroute = new Button { Text = "Trace結果クリア", Location = new Point(264, 22), Width = 140 }; // 名前変更
            btnSaveTraceroute = new Button { Text = "Trace結果保存", Location = new Point(412, 22), Width = 140 }; // 名前変更

            gbTrace.Controls.Add(btnTraceroute);
            gbTrace.Controls.Add(btnStopTraceroute);
            gbTrace.Controls.Add(btnClearTraceroute);
            gbTrace.Controls.Add(btnSaveTraceroute);

            // 左側にグループボックスを流す FlowLayoutPanel を作成して、それぞれの枠が"囲み切る"ように配置
            var leftFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                AutoSize = false,
                Width = gbPing.Width + gbTrace.Width + 24,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8),
                Margin = new Padding(0)
            };
            leftFlow.Controls.Add(gbPing);
            leftFlow.Controls.Add(gbTrace);

            // 右側に終了ボタンを置くためのパネル（枠に被らないように右端に固定）
            var rightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 120,
                Padding = new Padding(8)
            };

            // add exit button to the right (inside rightPanel)
            btnExit = new Button { Text = "終了", Anchor = AnchorStyles.Top | AnchorStyles.Right, Width = 80, Height = 30 };
            btnExit.Location = new Point(rightPanel.ClientSize.Width - btnExit.Width - 10, 20);
            // adjust location when resized
            rightPanel.Resize += (s, e) =>
            {
                btnExit.Location = new Point(Math.Max(8, rightPanel.ClientSize.Width - btnExit.Width - 10), 20);
            };
            rightPanel.Controls.Add(btnExit);

            // add flow and right panel to bottom panel
            bottomPanel.Controls.Add(rightPanel);
            bottomPanel.Controls.Add(leftFlow);

            // add to form
            if (this.MainMenuStrip != null && !this.Controls.Contains(this.MainMenuStrip))
            {
                this.Controls.Add(this.MainMenuStrip);
                this.MainMenuStrip.Dock = DockStyle.Top;
            }
            this.Controls.Add(contentPanel);
            this.Controls.Add(bottomPanel);

            // handlers
            btnPingStart.Click += BtnPingStart_Click;
            btnStop.Click += BtnStop_Click;
            btnClear.Click += BtnClear_Click;
            btnExit.Click += BtnExit_Click;
            btnSave.Click += BtnSaveResult_Click;
            btnTraceroute.Click += BtnTraceroute_Click;
            btnSaveTraceroute.Click += BtnSaveTraceroute_Click;
            btnClearTraceroute.Click += BtnClearTraceroute_Click;
            btnStopTraceroute.Click += BtnStopTraceroute_Click;

            // ensure menu spacing
            int menuHeight = 0;
            if (this.MainMenuStrip != null)
            {
                menuHeight = this.MainMenuStrip.PreferredSize.Height;
            }
            contentPanel.Padding = new Padding(0, menuHeight, 0, 0);

            // initial traceroute button state
            tracerouteTextBoxes = new Dictionary<string, TextBox>(StringComparer.OrdinalIgnoreCase);
            UpdateTracerouteButtons();
        }

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
            mnuAutoSaveAllPing.CheckedChanged += (s, e) => UpdateUiState(cts != null ? "Running" : (string.IsNullOrEmpty(txtStartTime.Text) ? "Initial" : "Stopped"));
            opt.DropDownItems.Add(mnuAutoSaveAllPing);
            mnuAutoSaveTraceroute = new ToolStripMenuItem("Trace結果の自動保存") { CheckOnClick = true, Checked = true };
            mnuAutoSaveTraceroute.CheckedChanged += (s, e) => UpdateTracerouteButtons();
            opt.DropDownItems.Add(mnuAutoSaveTraceroute);
            ms.Items.Add(opt);

            var help = new ToolStripMenuItem("ヘルプ");
            help.DropDownItems.Add("バージョン情報", null, VersionItem_Click);
            ms.Items.Add(help);

            this.MainMenuStrip = ms;
            this.Controls.Add(ms);

            // 重要: ここで BringToFront / SetChildIndex を行わない。
            // MenuStrip を後から追加することで、top-docked controls の下に配置され、
            // 期待するレイアウト（スクリーンショットの OK になる）になる。
        }

        private void SetupDataGridViewColumns()
        {
            dgvMonitor.Columns.Clear();
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "ステータス", HeaderText = "ｽﾃｰﾀｽ", Width = 60, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = "Trace", HeaderText = "Trace", Width = 60, TrueValue = true, FalseValue = false });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "順番", HeaderText = "順番", Width = 50, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "対象アドレス", HeaderText = "対象アドレス", Width = 120 });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Host名", HeaderText = "Host名", Width = 120 });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "送信回数", HeaderText = "送信回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗回数", HeaderText = "失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗回数", HeaderText = "連続失敗回数", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "連続失敗時間s", HeaderText = "連続失敗時間", Width = 130, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最大失敗時間s", HeaderText = "最大失敗時間", Width = 130, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "時間ms", HeaderText = "時間[ms]", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "平均値ms", HeaderText = "平均値[ms]", Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" }, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最小値ms", HeaderText = "最小値[ms]", Width = 80, ReadOnly = true });
            dgvMonitor.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "最大値ms", HeaderText = "最大値[ms]", Width = 80, ReadOnly = true });

            dgvMonitor.CellFormatting += (s, e) =>
            {
                if (e.ColumnIndex == 0 && dgvMonitor.Rows[e.RowIndex].DataBoundItem is PingMonitorItem item)
                {
                    if (item.ステータス == "Down") { e.CellStyle.BackColor = Color.MistyRose; e.CellStyle.SelectionBackColor = Color.Red; }
                    else if (item.ステータス == "OK") { e.CellStyle.BackColor = Color.LightGreen; e.CellStyle.SelectionBackColor = Color.Green; }
                    else if (item.ステータス == "復旧") { e.CellStyle.BackColor = Color.LightSkyBlue; e.CellStyle.SelectionBackColor = Color.Blue; }
                }
            };

            disruptionLogList = new SortableBindingList<DisruptionLogItem>(new List<DisruptionLogItem>());
            dgvLog.DataSource = disruptionLogList;
            dgvLog.ColumnHeaderMouseClick += dgvLog_ColumnHeaderMouseClick;
            dgvLog.Columns.Clear();
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "対象アドレス", HeaderText = "対象アドレス", Width = 120 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Host名", HeaderText = "Host名", Width = 120 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down開始日時", HeaderText = "Down開始日時", Width = 150 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧日時", HeaderText = "復旧日時", Width = 150 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗回数", HeaderText = "失敗回数", Width = 80 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "失敗時間mmss", HeaderText = "失敗時間", Width = 100 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前平均ms", HeaderText = "Down前平均", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最小ms", HeaderText = "Down前最小", Width = 110 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Down前最大ms", HeaderText = "Down前最大", Width = 110 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後平均ms", HeaderText = "復旧後平均", Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "F1" } });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最小ms", HeaderText = "復旧後最小", Width = 110 });
            dgvLog.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "復旧後最大ms", HeaderText = "復旧後最大", Width = 110 });
        }
    }
}